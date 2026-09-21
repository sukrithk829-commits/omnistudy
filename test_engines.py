import unittest
import os
import io
import json
import sqlite3
import tempfile
from unittest.mock import patch, MagicMock

# Import app modules
import app
from app import (
    clean_study_text, extract_clean_sentences, summarize, make_quiz,
    build_smart_options, detect_domain, _parse_gemini_json, _valid_study_pack,
    build_flashcards, extract_concept_graph, generate_cheat_sheet_data,
    generate_viva_questions, evaluate_viva_answer, comparison_rows,
    visible_document, validate_document_access, allowed_file, extract_upload
)

class ComprehensiveOmniStudyScan(unittest.TestCase):
    def setUp(self):
        app.app.config['TESTING'] = True
        app.app.config['WTF_CSRF_ENABLED'] = False
        self.client = app.app.test_client()

    def test_clean_study_text(self):
        sample = "Page 1\nTABLE OF CONTENTS\n■ A Complete Guide to Foundations\nRole: Filters blood\nWatch out for: High BP"
        cleaned = clean_study_text(sample)
        self.assertIsInstance(cleaned, str)

    def test_extract_clean_sentences(self):
        sample = "The heart pumps oxygenated blood through the body. Mitochondria is the powerhouse of the cell."
        sentences = extract_clean_sentences(sample)
        self.assertIsInstance(sentences, list)

    def test_summarize_empty(self):
        self.assertEqual(summarize(""), [])
        self.assertEqual(summarize("Short"), ["Short"])

    def test_make_quiz_generation(self):
        content = """
        Heart
        Role: Pumps oxygenated blood through the arterial network.
        Watch out for: Arrhythmia and coronary arterial blockage.
        
        Brain
        Role: Coordinates neurological signals and systemic homeostasis.
        Watch out for: Ischemic stroke and neurological deficits.
        
        Vitamin C
        Role: Synthesizes collagen and neutralizes reactive oxygen species.
        
        Algorithms — Systematic procedures for computational problem solving.
        
        A Binary Search Tree is a hierarchical structure that optimizes search operations.
        """
        quiz = make_quiz(content, max_questions=10)
        self.assertIsInstance(quiz, list)
        for q in quiz:
            self.assertIn("prompt", q)
            self.assertIn("options", q)
            self.assertIn("answer", q)
            self.assertEqual(len(q["options"]), 4)
            self.assertIn(q["answer"], q["options"])

    def test_build_smart_options_deduplication(self):
        correct = "Primary Option"
        candidates = ["Primary Option", "primary option", "Candidate 2", "Candidate 3"]
        fallbacks = ["Fallback 1", "Fallback 2"]
        opts = build_smart_options(correct, candidates, fallbacks)
        self.assertEqual(len(opts), 4)
        self.assertEqual(len(set(opts)), 4)
        self.assertIn(correct, opts)

    def test_gemini_json_parsing(self):
        fenced_json = "```json\n{\"summary\": [\"P1\"], \"questions\": []}\n```"
        parsed = _parse_gemini_json(fenced_json)
        self.assertEqual(parsed["summary"], ["P1"])

        raw_json = "{\"summary\": [\"P2\"], \"questions\": []}"
        parsed2 = _parse_gemini_json(raw_json)
        self.assertEqual(parsed2["summary"], ["P2"])

    def test_valid_study_pack(self):
        pack = {
            "summary": ["Point 1"],
            "questions": [
                {
                    "prompt": "What is X?",
                    "options": ["A", "B", "C", "D"],
                    "answer": "A",
                    "explanation": "Because A",
                    "skill": "Concept",
                    "difficulty": "Easy"
                }
            ]
        }
        res = _valid_study_pack(pack)
        self.assertIsNotNone(res)
        self.assertEqual(len(res["summary"]), 1)
        self.assertEqual(len(res["questions"]), 1)

    def test_flashcards_builder(self):
        content = "Heart\nRole: Pumps blood\nWatch out for: Chest pain\nB-Tree — Balanced search tree\nThe cache optimizes memory latency."
        cards = build_flashcards(content)
        self.assertTrue(len(cards) > 0)
        for c in cards:
            self.assertIn("id", c)
            self.assertIn("front", c)
            self.assertIn("back", c)
            self.assertIn("category", c)

    def test_concept_graph(self):
        content = "Module 1\nDatabase — A structured collection of data\nThe DBMS manages transaction serializability."
        graph = extract_concept_graph(content, doc_title="Database Management Systems")
        self.assertIn("nodes", graph)
        self.assertIn("edges", graph)
        self.assertIn("metrics", graph)

    def test_cheat_sheet(self):
        content = "Heart\nRole: Pumps blood\nAlgorithm — Step by step procedure"
        cs = generate_cheat_sheet_data(content)
        self.assertIn("definitions", cs)
        self.assertIn("mechanisms", cs)
        self.assertIn("exam_qa", cs)

    def test_viva_evaluation(self):
        eval_res = evaluate_viva_answer(["blood", "heart", "oxygen"], "The heart pumps oxygenated blood", "The heart pumps oxygen and blood to tissues")
        self.assertIn("score", eval_res)
        self.assertIn("grade", eval_res)
        self.assertIn("feedback", eval_res)
        self.assertGreaterEqual(eval_res["score"], 2.0)

    def test_comparison_rows(self):
        rows = comparison_rows("Line 1\nLine 2", "Line 1\nLine 2 modified\nLine 3")
        self.assertIsInstance(rows, list)
        self.assertTrue(len(rows) >= 2)

    def test_quiz_cs_distractors_and_stop_words(self):
        content = """
        Database System
        A database is an organized collection of related data stored electronically.
        A Data Model is a collection of conceptual tools used to describe data structure.
        Primary Key: A unique attribute that unambiguously identifies a tuple in a table.
        Foreign Key: A field that refers to the primary key in another table.
        Both are paired tags in HTML.
        Both tags support cite and datetime attributes.
        """
        quiz = make_quiz(content, max_questions=15)
        self.assertTrue(len(quiz) >= 3)
        for q in quiz:
            # Ensure "Both" is not treated as a valid concept
            self.assertNotEqual(q.get("answer"), "Both")
            self.assertFalse(q["prompt"].startswith("What is the primary action, mechanism, or principle associated with \"Both\""))
            self.assertEqual(len(q["options"]), 4)
            # Ensure punctuation matching: if answer ends with period, all options end with period
            if q["answer"].endswith("."):
                for opt in q["options"]:
                    self.assertTrue(opt.endswith("."))

    def test_select_quiz_questions_seeded_rotation(self):
        from app import select_quiz_questions
        # Create a mock bank of 20 questions
        bank = [
            {"prompt": f"Prompt {i}", "options": [f"Opt {i}_A", f"Opt {i}_B", f"Opt {i}_C", f"Opt {i}_D"], "answer": f"Opt {i}_A", "difficulty": "Easy" if i % 2 == 0 else "Medium"}
            for i in range(20)
        ]
        # Same seed produces exact same selection
        sel1 = select_quiz_questions(bank, "balanced", 5, seed="seed_123")
        sel2 = select_quiz_questions(bank, "balanced", 5, seed="seed_123")
        self.assertEqual([q["prompt"] for q in sel1], [q["prompt"] for q in sel2])
        self.assertEqual([q["options"] for q in sel1], [q["options"] for q in sel2])

        # Different seed produces different selection (Try another set)
        sel3 = select_quiz_questions(bank, "balanced", 5, seed="seed_999")
        # At least one prompt or option order should differ across different seeds
        prompts1 = [q["prompt"] for q in sel1]
        prompts3 = [q["prompt"] for q in sel3]
        self.assertNotEqual(prompts1, prompts3)


if __name__ == '__main__':
    unittest.main()
