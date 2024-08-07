

import logging
from xlcalculator import ModelCompiler
from xlcalculator import Model
from xlcalculator import Evaluator
import os
import pprint 
logging.basicConfig(level=logging.DEBUG)
import time
import sys

sys.setrecursionlimit(15000)


#Sheet1!B500 125250
# Time taken to evaluate 'Sheet1!B49': 0.011705875396728516 seconds
# json_file_name = r'./examples/joel_test1/test1.json'
# Sheet1!B500 125250
# Time taken to evaluate 'Sheet1!B49': 0.026947975158691406 seconds

# filename = r'./examples/joel_test1/test1.xlsx'
# filename = r'./examples/joel_test1/oracle_building_1.xlsx'
# filename = r'./oracle_building_1.xlsx'
filename = r'./calc_tree.xlsx'
compiler = ModelCompiler()
print(f" ###### Current working directory: {os.getcwd()}")
new_model = compiler.read_and_parse_archive(filename, build_code=True)
# new_model.persist_to_json_file(json_file_name)
# reconstituted_model = Model()
# reconstituted_model.construct_from_json_file(json_file_name, build_code=True)

evaluator = Evaluator(new_model, None, freeze_cells_mode=True)

start_time = time.time()

sheet1 = 'Sheet1!B500'
m2 = evaluator.evaluate(sheet1)
print(sheet1, m2)

sheet1 = 'Sheet1!B501'
m2 = evaluator.evaluate(sheet1)
print(sheet1, m2)


end_time = time.time()

time1 = end_time - start_time
print(f"Time taken to evaluate 1: {time1:10f} seconds")

# val20 = evaluator.set_cell_value('Sheet1!C1',99)

start_time2 = time.time()

sheet1 = 'Sheet1!B500'
m2 = evaluator.evaluate(sheet1)
print(sheet1, m2)

sheet1 = 'Sheet1!B501'
m2 = evaluator.evaluate(sheet1)
print(sheet1, m2)


end_time2 = time.time()

time2 = end_time2 - start_time2
print(f"Time taken to evaluate 2: {time2:10f} seconds")


