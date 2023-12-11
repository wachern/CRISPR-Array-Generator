"""
Test cases for crispr_array_generator
Created: Fall 2023
Author: Willow Chernoske
"""

import unittest
import os
import os.path
from crispr_array_generator.crisprarraygenerator import Array

class TestArray(unittest.TestCase):

    def setUp(self):
        # Forward and reverse DNA strands
        self.forward = 'atgcgga'
        self.reverse = 'tccgcat'

        # TCC and/or non-nucleotide grnas removed
        self.grnas_input = ['ttcaaaggg', 'cccccc', 'tchtaa']
        self.grnas_output1 = ['aaaggg', 'cccccc']
        self.grnas_output2 = ['ttcaaaggg', 'cccccc']

    def test_get_reverse_complement(self):
        # Function outputs valid DNA reverse compliment
        result = Array.get_reverse_complement(self.forward)
        self.assertIsNotNone(result)
        self.assertEqual(result, self.reverse)

    def test_extract_excel_data(self):
        # Function can access and correctly process DNA data from an excel file
        path = os.path.join(os.getcwd(), "testfile.xlsx")
        self.assertTrue(os.path.isfile(path))
        result = Array.extract_excel_data('testfile')
        self.assertEqual(result, self.grnas_output2)

    def test_check_grna_array_input(self):
        # Function removes TTC PAM sequence and non-nucleotide characters
        result = Array.check_grna(self.grnas_input)
        self.assertIsNotNone(result)
        self.assertEqual(result, self.grnas_output1)
        # Function creates an excel file "array_report.xlsx"
        path = os.path.join(os.getcwd(), "array_report.xlsx")
        self.assertTrue(os.path.isfile(path))
        if os.path.isfile(path):
            os.remove(path)

    def test_check_grna_excel_input(self):
        # Function removes TTC PAM sequence and non-nucleotide characters
        result = Array.check_grna('testfile')
        self.assertIsNotNone(result)
        self.assertEqual(result, self.grnas_output1)
        # Function creates an excel file "array_report.xlsx"
        path = os.path.join(os.getcwd(), "array_report.xlsx")
        self.assertTrue(os.path.isfile(path))
        if os.path.isfile(path):
            os.remove(path)
