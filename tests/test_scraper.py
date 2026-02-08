"""
Unit tests for the scraper module
"""

import unittest
import sys
from pathlib import Path

# Add src to path
src_path = Path(__file__).parent.parent / 'src'
sys.path.insert(0, str(src_path))

from scraper import get_engineer_syndicate_safe


class TestScraper(unittest.TestCase):
    """Test cases for scraper functionality"""
    
    def test_invalid_input_type(self):
        """Test that non-string input is rejected"""
        result = get_engineer_syndicate_safe(12345678901234)
        self.assertFalse(result['success'])
        self.assertIn('Validation Error', result['error'])
    
    def test_invalid_length(self):
        """Test that wrong length IDs are rejected"""
        result = get_engineer_syndicate_safe("123456789")
        self.assertFalse(result['success'])
        self.assertIn('14 digits', result['error'])
    
    def test_invalid_characters(self):
        """Test that non-digit characters are rejected"""
        result = get_engineer_syndicate_safe("1234567890123a")
        self.assertFalse(result['success'])
        self.assertIn('14 digits', result['error'])
    
    def test_valid_format(self):
        """Test that valid format passes validation"""
        # Note: This test may fail if the ID doesn't exist in the database
        # or if there's a network issue
        result = get_engineer_syndicate_safe("29501011234567")
        # We only check that it doesn't fail on validation
        if not result['success']:
            # If it fails, it should be a data error, not validation
            self.assertNotIn('Validation Error', result['error'])
    
    def test_result_structure(self):
        """Test that result has the correct structure"""
        result = get_engineer_syndicate_safe("29501011234567")
        self.assertIn('success', result)
        self.assertIn('national_id', result)
        if result['success']:
            self.assertIn('syndicate', result)
        else:
            self.assertIn('error', result)


class TestScraperIntegration(unittest.TestCase):
    """Integration tests that require network access"""
    
    @unittest.skip("Requires network access and valid test ID")
    def test_real_lookup(self):
        """Test actual lookup with a real ID (requires network)"""
        # Replace with a real test ID if available
        result = get_engineer_syndicate_safe("XXXXXXXXXXXXXX")
        self.assertTrue(result['success'])
        self.assertIn('syndicate', result)


if __name__ == '__main__':
    unittest.main()