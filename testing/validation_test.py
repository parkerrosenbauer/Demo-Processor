import unittest
import datetime
from file_processing.validation import Validation
from file_processing.demo import Demo


class TestValidation(unittest.TestCase):
    def test_retrieve_one(self):
        this_demo = Validation(datetime.datetime.strptime("10/6/2022", '%m/%d/%Y'))
        self.assertEqual(0, this_demo.retrieve_one("flipped_open"))
        self.assertRaises(KeyError, this_demo.retrieve_one, "bad_count")

    def test_retrieve_all(self):
        that_demo = Validation(datetime.datetime.strptime("9/20/2022", '%m/%d/%Y'))
        counts = that_demo.retrieve_all()
        self.assertEqual(0, counts["contact_no_lead"])
        self.assertEqual([], counts["udb_tracking_codes"])

        not_real_demo = Validation(datetime.datetime.strptime("9/30/2022", '%m/%d/%Y'))
        self.assertRaises(KeyError, not_real_demo.retrieve_all)

    def test_update_counts(self):
        another_demo = Validation(datetime.datetime.strptime("10/5/2022", '%m/%d/%Y'))
        track = ["UAC1", "UBC1"]
        another_demo.update_counts(flipped_open=5, udb_tracking_codes=track)
        self.assertEqual(5, another_demo.retrieve_one("flipped_open"))
        self.assertEqual(track, another_demo.retrieve_one("udb_tracking_codes"))
        self.assertRaisesRegex(ValueError, "bad_count is not a valid metric.", another_demo.update_counts, bad_count=2)

    def test_with_demo(self):
        a_demo = Demo(datetime.datetime.strptime("10/5/2022", '%m/%d/%Y'))
        counts = a_demo.counts.retrieve_all()
        self.assertEqual(0, counts["flipped_open"])
