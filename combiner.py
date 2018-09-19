import sys 
import os
import glob
import openpyxl

def read_files(dir): 
    '''
        reads the data of all the csv files in a given directory into a list
    '''
    filenames = glob.glob(os.path.join(dir, '*.csv')) 
    return [read_data(filename) for filename in filenames] 


def read_data(filename):
    '''
        reads the first 4 elements of the second line of the file and creates the time period element 
    '''
    with open(filename,'r') as f:
            data_line = f.readlines()[1]
            req_data = data_line.split(',')[:4]
            req_data.append(determine_time_period(req_data[0]))
            return req_data

def determine_time_period(date_time): 
    '''
        return '5pm' if the time ends with a PM i.e. time is equal or after 12. else, returns '1am'
    '''
    return '5pm' if date_time.endswith('PM') else "1am"

def write(data, output_file, sheet_name='MP Location Utilisations'): 
    '''
        writes the data into the file
    '''
    workbook = openpyxl.load_workbook(output_file)
    sheet = workbook[sheet_name]
    for row in data: 
        sheet.append(row) 
    workbook.save(output_file) 

def is_args_valid(input_dir, output_file):
    '''
        checks if the input directory is a valid dir and if the output file is a valid file 
    '''
    return os.path.isdir(input_dir) and os.path.isfile(output_file) 

def main(): 
    '''
        the starting point of the program 
        conventionally named 'main' but can be anything else 
    '''
    input_dir = os.path.abspath('input')
    output_file = os.path.abspath('output/output.xlsx')

    if len(sys.argv) == 3: 
        input_dir = sys.argv[1]
        output_file = sys.argv[2] 

    if not is_args_valid(input_dir, output_file): 
        print("ERROR: Invalid input directory or output file.")
        return

    data_to_write = read_files(input_dir)  
    write(data_to_write, output_file)   

if __name__ == '__main__': 
    main()
