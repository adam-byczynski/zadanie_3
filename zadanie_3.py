import time
import os
import datetime
import random
import pandas
from functools import wraps
from collections import namedtuple


Row = namedtuple('DataRow', ['label', 'insertion', 'selection', 'merge'])


class DataRow:
    rows = []
    label = None
    insertion = None
    selection = None
    merge = None
    insertion_average = None
    selection_average = None
    merge_average = None
    insertion_std = None
    selection_std = None
    merge_std = None


class DataGenerator():
    __test_4_add_parameters = (1, 500, 400)

    @staticmethod
    def random(count):
        return DataGenerator.generate_list_of_random_numbers(count)

    @staticmethod
    def sorted(count):
        temp_list = DataGenerator.generate_list_of_random_numbers(count)
        return sorted(temp_list)

    @staticmethod
    def reverse_sorted(count):
        temp_list = DataGenerator.generate_list_of_random_numbers(count)
        return sorted(temp_list, reverse=True)

    @staticmethod
    def sorted_with_addition(count):
        temp_list = DataGenerator.generate_list_of_random_numbers(count)
        (start, end, how_many) = DataGenerator.__test_4_add_parameters
        return sorted(temp_list) + random.sample(range(start, end), how_many)

    @staticmethod
    def generate_list_of_random_numbers(count):
        range_start = 1
        range_end = count * 2
        list_size = count
        return random.sample(range(range_start, range_end), list_size)


class Algorithms():
    def algorithm_timer(func):
        @wraps(func)
        def timer(*args, **kwargs):
            time_start = time.time()
            algorithm = func(*args, **kwargs)
            time_elapsed = time.time() - time_start
            formatted_time_elapsed = format(time_elapsed, '.10f')
            kwargs['measurements'].append(float(formatted_time_elapsed))
            return algorithm
        return timer

    @algorithm_timer
    def insertion_sort(test_data, measurements=[]):
        numbers = test_data.copy()
        for i in range(1, len(numbers)):
            value = numbers[i]
            j = i - 1
            while j >= 0 and value < numbers[j]:
                numbers[j + 1] = numbers[j]
                j -= 1
            numbers[j + 1] = value

    @algorithm_timer
    def selection_sort(test_data, measurements=[]):
        numbers = test_data.copy()
        for number_index in range(len(numbers) - 1):
            current_sublist = numbers[number_index:]
            min_value = min(current_sublist)
            relative_min_index = current_sublist.index(min_value)
            numbers.insert(number_index, numbers.pop(relative_min_index + number_index))

    @algorithm_timer
    def merge_sort(test_data, measurements=[]):
        numbers = test_data.copy()

        def mergeSort(myList):
            if len(myList) > 1:
                mid = len(myList) // 2
                left = myList[:mid]
                right = myList[mid:]

                mergeSort(left)
                mergeSort(right)

                i = 0
                j = 0
                k = 0

                while i < len(left) and j < len(right):
                    if left[i] < right[j]:
                        myList[k] = left[i]
                        i += 1
                    else:
                        myList[k] = right[j]
                        j += 1
                    k += 1

                while i < len(left):
                    myList[k] = left[i]
                    i += 1
                    k += 1

                while j < len(right):
                    myList[k] = right[j]
                    j += 1
                    k += 1
        mergeSort(numbers)


class Tester():
    def __init__(self):
        self.insertion_results = []
        self.selection_results = []
        self.merge_results = []
        self.results = []

    def run_full_tests(self, data_generator, number_of_tests=5, number_of_subtests=10,
                           data_count=2500, count_increment=2500):
        current_test = 1
        count = data_count
        while current_test <= number_of_tests:
            current_subtest = 1
            while current_subtest <= number_of_subtests:
                test_data = data_generator(count)
                self.single_subtest(test_data)
                current_subtest += 1
            self.results.append(self.build_result(current_test, number_of_subtests, count))
            self.clear_partial_results()
            current_test += 1
            count += count_increment

    def single_subtest(self, test_data):
        Algorithms.insertion_sort(test_data, measurements=self.insertion_results)
        Algorithms.selection_sort(test_data, measurements=self.selection_results)
        Algorithms.merge_sort(test_data, measurements=self.merge_results)

    def build_result(self, current_test, number_of_subtests, count):
        result = DataRow()
        result.count = count
        result.rows = []
        for current_subtest in range(number_of_subtests):
            label = f'{str(current_test)}-{str(current_subtest + 1)}'
            result.rows.append(Row(label, self.insertion_results[current_subtest], self.selection_results[current_subtest], self.merge_results[current_subtest]))
        from statistics import mean
        result.insertion_average = mean(self.insertion_results)
        result.selection_average = mean(self.selection_results)
        result.merge_average = mean(self.merge_results)
        from statistics import pstdev
        result.insertion_std = pstdev(self.insertion_results)
        result.selection_std = pstdev(self.selection_results)
        result.merge_std = pstdev(self.merge_results)
        return result
        
    def clear_partial_results(self):
        self.insertion_results.clear()
        self.selection_results.clear()
        self.merge_results.clear()

    def clear_all_results(self):
        self.insertion_results.clear()
        self.selection_results.clear()
        self.merge_results.clear()
        self.results.clear()
        

class ExcelExporter():
    DataSheet = namedtuple('DataSheet', ['sheet_name', 'data_frame'])

    def __init__(self):
        self.filename = ExcelExporter.generate_filename()
        self.data_sheets = []

    @staticmethod
    def generate_filename():
        formatted_current_datetime = datetime.datetime.now().strftime("%y-%m-%d_%H_%M")
        return f"Sorting Algorithms Measurements_{formatted_current_datetime}"

    def export_file(self):
        with pandas.ExcelWriter(f"{self.filename}.xlsx") as file:
            for sheet_name, data_frame in self.data_sheets:
                data_frame.to_excel(file, sheet_name=f'{sheet_name}')

    def launch_file(self):
        os.system(f"start EXCEL.EXE \"{self.filename}.xlsx\"")

    def generate_sheet(self, data, sheet_name):
        data_frame = ExcelExporter.create_data_frame(data)
        self.data_sheets.append(self.DataSheet(sheet_name, data_frame))

    @staticmethod
    def create_data_frame(data):
        def pad_with_none(arr, target_length):
            return arr + [None] * (target_length - len(arr))

        test_numbers = []
        numbers_count = []
        insertion_results = []
        insertion_averages = []
        insertion_deviations = []
        selection_results = []
        selection_averages = []
        selection_deviations = []
        merge_results = []
        merge_averages = []
        merge_deviations = []
        for item in data:
            rows = len(item.rows)
            test_numbers.extend([row.label for row in item.rows])
            numbers_count.extend(pad_with_none([item.count], rows))
            insertion_results.extend([row.insertion for row in item.rows])
            insertion_averages.extend(pad_with_none([item.insertion_average], rows))
            insertion_deviations.extend(pad_with_none([item.insertion_std], rows))
            selection_results.extend([row.selection for row in item.rows])
            selection_averages.extend(pad_with_none([item.selection_average], rows))
            selection_deviations.extend(pad_with_none([item.selection_std], rows))
            merge_results.extend([row.merge for row in item.rows])
            merge_averages.extend(pad_with_none([item.merge_average], rows))
            merge_deviations.extend(pad_with_none([item.merge_std], rows))
            
        measurements_data = {
            'Test number': test_numbers,
            'Numbers count': numbers_count,
            'Insertion sort': insertion_results,
            'Insertion Avg': insertion_averages,
            'Insertion SD': insertion_deviations,
            'Selection sort': selection_results,
            'Selection Avg': selection_averages,
            'Selection SD': selection_deviations,
            'Merge sort': merge_results,
            'Merge Avg': merge_averages,
            'Merge SD': merge_deviations,
        }

        measurements_data_frame = pandas.DataFrame(measurements_data, dtype=float)
        measurements_data_frame = measurements_data_frame.set_index('Test number')
        return measurements_data_frame


RunConfig = namedtuple('RunConfig', ['sheet_name', 'generator'])


def main():
    run_configs = (
        RunConfig("Random order", DataGenerator.random), 
        RunConfig("Sorted", DataGenerator.sorted),
        RunConfig("Reverse-sorted", DataGenerator.reverse_sorted), 
        RunConfig("Sorted+addition", DataGenerator.sorted_with_addition)
    )

    exporter = ExcelExporter()
    for config in run_configs:
        tester = Tester()
        tester.run_full_tests(config.generator)
        exporter.generate_sheet(tester.results, config.sheet_name)
        tester.clear_all_results()
    exporter.export_file()
    exporter.launch_file()


if __name__ == '__main__':
    main()
    