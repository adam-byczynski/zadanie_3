import unittest
import unittest.mock
import zadanie_3_bambo as zad3
import io
import random


class TestSortingAlgorithms(unittest.TestCase):
    def generate_list_of_random_numbers(range_start=1, range_end=100, list_size=10):
        return random.sample(range(range_start, range_end), list_size)

    generated_nums = generate_list_of_random_numbers()

    def test_insertion_sort(self):
        test_list = self.generated_nums
        expected = test_list.copy()
        expected.sort()
        result = zad3.Algorithms.insertion_sort(test_list.copy())
        self.assertEqual(expected, result)

    def test_selection_sort(self):
        test_list = self.generated_nums
        expected = test_list.copy()
        expected.sort()
        result = zad3.Algorithms.selection_sort(test_list.copy())
        self.assertEqual(expected, result)

    def test_merge_sort(self):
        test_list = self.generated_nums
        expected = test_list.copy()
        expected.sort()
        result = zad3.Algorithms.merge_sort(test_list.copy())
        self.assertEqual(expected, result)


if __name__ == '__main__':
    unittest.main()
