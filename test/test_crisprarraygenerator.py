"""
Test cases for crispr_array_generator
Created: Fall 2023
Author: Willow Chernoske
"""

import unittest
from crispr_array_generator.crisprarraygenerator import Array

class TestArray(unittest.TestCase):

    def setUp(self):
        #forward and reverse
        self.forward = 'atgcgga'
        self.reverse = 'tccgcat'

        #removes TCC and non-nucleotide grnas
        self.grnas_input = ['ttcaaaggg', 'cccccc', 'tchtaa']
        self.grnas_output = ['aaaggg', 'cccccc']

    def test_get_reverse_complement(self):
        result = Array.get_reverse_complement(self.forward)
        self.assertIsNotNone(result)
        self.assertEqual(result, self.reverse)
        print("OK!")

    def test_check_grna(self):
        result = Array.check_grna(self.grnas_input)
        self.assertIsNotNone(result)
        self.assertEqual(result, self.grnas_output)
        print("OK2!")
