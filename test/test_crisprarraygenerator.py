"""
Created: Fall 2023
Author: Willow Chernoske
"""

import unitest
from crispr_array_generator.crisprarraygenerator import Array

class TestArray(unittest.TestCase):

    def setUp(self):
        forward = atgcgga
        reverse = tccgcat

    def test_get_reverse_complement(self)
        result = self.forward.get_reverse_complement()
        self.assertIsNotNone(result)
        self.assertEqual(reverse)
