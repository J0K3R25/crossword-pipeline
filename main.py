import json
import puz
import re

# prepare google sheets api
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

sheet_id = "1RTbr5EtiZViUXQvDyiqf-ksV-OqcPBW5-kHbRW38QUU"
sh = client.open_by_key(sheet_id)
service = build('sheets', 'v4', credentials=creds)


# prepare automatic solver
from solver import Utils 
from solver.Crossword import Crossword
from solver.BPSolver import BPSolver
import sys
import os
# Add the DPR folder to the system path
sys.path.append(os.path.join(os.path.dirname(__file__), 'DPR'))

def solve(crossword):
    solver = BPSolver(crossword, max_candidates=500000)
    solution = solver.solve(num_iters=10, iterative_improvement_steps=5)
    print("*** Solver Output ***")
    Utils.print_grid(solution)

    return solution

# convert the column index to the according letter
def col_num_to_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result  # 65 is the ASCII value for 'A'
        n = n // 26 - 1
    return result

# takes the coordinates of a cell and returns the letter representation, e.g. 0,0 becomes A1
def def_cell(row1, col1):
    return (col_num_to_letter(col1) + str(row1+1))


# takes the coordinates of two cells and returns them in cell range format, e.g. 0,0,1,0 becomes A1:A2
def def_cell_range(row1, col1, row2, col2):
    return (def_cell(row1,col1)+ ":" + def_cell(row2,col2))

# uses threshold to determine if cell is white 
def is_white_cell(r,g,b,threshold):
    # Convert to grayscale for color thresholding
    grayscale = 0.299 * r + 0.587 * g + 0.114 * b
    return grayscale > threshold 

# return true if the background of a certain cell set, return false if it is white
# imput is a cell in sheet cell format, like "A1"
def is_background_set_single_cell(cell):
    # Send the request to get background color information
    request = service.spreadsheets().get(
        spreadsheetId=sheet_id,
        ranges=f'{sheet_name}!{cell}',
        fields="sheets.data.rowData.values.effectiveFormat.backgroundColor"
    )
    response = request.execute()
    #print(f"Response for cell {cell}: {response}")  # Debug: Print the response

    try:
        # Access the background color directly from the response
        background_color = response['sheets'][0]['data'][0]['rowData'][0]['values'][0]['effectiveFormat']['backgroundColor']
        
        # Check if the color is white (RGB: 1, 1, 1)
        red = background_color.get('red', 0)
        green = background_color.get('green', 0)
        blue = background_color.get('blue', 0)

        if (is_white_cell(red,green,blue,0.3)):
            return False  # Background is white
        else:
            return True  # Background is set to something else
    except KeyError:
        return True  # Assume background is set to something else

def is_background_set(range):
    # Send request to get background color information for the entire range (the whole crossword grid) 
    request = service.spreadsheets().get(
        spreadsheetId=sheet_id,
        ranges=f'{sheet_name}!{range}',
        fields="sheets.data.rowData.values.effectiveFormat.backgroundColor"
    )
    response = request.execute()

    # Initialize a result list for background states
    background_states = []

    try:
        rows_data = response['sheets'][0]['data'][0]['rowData']
        for row_data in rows_data:
            row_states = []
            for cell_data in row_data.get('values', []):
                # Get the effective background color
                background_color = cell_data.get('effectiveFormat', {}).get('backgroundColor', {})
                
                # Check if the color is white (RGB: 1, 1, 1)
                red = background_color.get('red', 0)
                green = background_color.get('green', 0)
                blue = background_color.get('blue', 0)

                
                # Check if the background is set to something other than white
                if (is_white_cell(red,green,blue,0.3)):
                    row_states.append(False)  # Background is white
                else:
                    row_states.append(True)   # Background is set to something else
            background_states.append(row_states)

    except KeyError:
        print("Key error with " + cell_data)
        return [[False] * len(rows_data[0]['values'])] * len(rows_data)  # Assuming all backgrounds are white if KeyError

    return background_states

# Checks if each word in the crossword contains at least 3 letters 
def is_valid_crossword(background_states):
    
    for row in background_states:
        if not check_line(row):
            return False

    for col_idx in range(size_crossword):
        column = [background_states[row_idx][col_idx] for row_idx in range(size_crossword)]
        if not check_line(column):
            return False

    return True

# Returns false if a line (row/col) contains a word with less than 3 letters and true otherwise 
def check_line(line):
    n = len(line)
    i = 0

    while i < n:
        if line[i]:  # Black cell
            if i + 1 < n and line[i + 1]:  # Next cell is also black
                i += 1
            elif i + 1 == n:  # End of the line
                return True
            else:
                # Check for at least 3 consecutive white cells
                white_count = 0
                while i + 1 < n and not line[i + 1]:
                    white_count += 1
                    i += 1
                if white_count < 3:
                    return False
        else:
            i += 1

    return True

# Determines size of the crossword
def auto_determine_size():
    first_row = sheet.row_values(1)
    try:
        col_index = first_row.index("Across") - 1  # Adding 1 because gspread uses 1-based indexing for columns
        return col_index
    except ValueError:
        print("Couldn't automatically determine the size.")
        return ask_size()

# Sets sheet to the one specified by the user
def ask_sheet_nr():
    while True:
        sheet_nr = input("Which Sheet is being used (e.g. type 1 for the first sheet) \n")
        
        if sheet_nr.isdigit() and int(sheet_nr) > 0:
            sheet_nr = int(sheet_nr)
            break 
        else:
            print("Please enter a valid number.")
    
    # Retrieve sheet metadata and get name of worksheet
    spreadsheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    sheet_properties = sheets[sheet_nr - 1].get('properties')  # Adjust for the sheet_nr (index + 1)
    sheet_name = sheet_properties.get('title')  # Get the sheet name dynamically

    return sh.get_worksheet(sheet_nr-1) , sheet_name

# Prints solution to Sheet if user answers with 'y'
def ask_print():
    while True:
        action = input("Do you want to fill the result into the sheet? y/n \n")
        
        # Check if input is valid
        if action == 'y':
            sheet.update(solution, def_cell_range(0,0,size_crossword,size_crossword))
            break
        elif action == 'n':
            break
        else:
            print("Please type y for yes and n for no.")

# Initialize crossword dictionary
def create_crossword_dict():
    crossword_dict =  {
        'metadata': {
            'date': None,   
            'rows': size_crossword,  
            'cols': size_crossword   
        },
        'clues': {
            'across': {},  
            'down': {}     
        },
        'grid': [] 
    }
    
    return crossword_dict

# Uses regular expression to get number and clue from the value
def process_clue(clue):
    #  Match a number at the start followed by any characters
    try:
        clue = clue.strip()
        match = re.match(r'^(\d+)[^\w]*(.*)', clue)
        if match:
            number = match.group(1)  # The numeric part
            clue = match.group(2)  # The clue part
        else:
            raise ValueError("Failed to match clue to regular expression.")
        return number,clue
    except Exception as e:
        print(f"Failed to process the clue: '{clue}'. Error: {e}")
        
# extracts the number for the length from the given value
def process_lenght(leng):
    trash,lenght = leng.split('(',1)
    lenght = lenght[:-1]
    return int(lenght)

# Creates a string of 'A's with specified length
def generate_answers(lenghts):
    answers = []
    for len in lenghts:
        answer = 'A' * len
        answers.append(answer)
    return answers

# Reads the clues of the specified column and generates according answers consisting of only letter 'A'
def read_col(column_nr):
    col = sheet.col_values(column_nr)
    col_processed = {}

    assert(len(col)>0)
    del(col[0])

    col_lenght = sheet.col_values(column_nr+1)
    del(col_lenght[0])
    lenghts = []

    for leng in col_lenght:
        lenghts.append(process_lenght(leng))

    answers = generate_answers(lenghts)
    
    for i in range(len(col)):
        number,clue = process_clue(col[i])
        col_processed[number] = [clue,answers[i]]

    return col_processed

def read_across():
    return read_col(size_crossword+2)

def read_down():
    return read_col(size_crossword+5)

def create_grid(size_crossword):
    # initialize every cell as 'BLACK'
    return [['BLACK' for _ in range(size_crossword)] for _ in range(size_crossword)]
    
# Creates a grid with black and white cells according to sheet and numbers white cells
def read_grid(grid):
    background_states = is_background_set(def_cell_range(0,0,size_crossword-1,size_crossword-1))
    if(not is_valid_crossword(background_states)):
        print("The crossword you provided might contain cells which are not crossed by two words. Make sure this is not the case.")
        input("Press enter to continue anyways")

    clue_number = 1 
    rows, cols = size_crossword, size_crossword

    for row in range(rows):
        for col in range(cols):
            # Check if the current cell is a black cell
            if background_states[row][col]:
                grid[row][col] = 'BLACK'
            else:
                needs_number = False
                
                # Check if it's the first cell in the row or if the left cell is black
                if col == 0:
                    needs_number = True
                elif background_states[row][col - 1]:  # Check left cell
                    needs_number = True
                
                # Check if it's the first cell in the column or if the above cell is black
                if row == 0:
                    needs_number = True
                elif background_states[row - 1][col]:  # Check above cell
                    needs_number = True
                
                # Assign clue number if necessary
                if needs_number:
                    grid[row][col] = [str(clue_number), 'A']
                    clue_number += 1
                else:
                    grid[row][col] = ['', 'A']

# Fill the crossword dictionary with the clues from the sheet
def fill_crossword_dict(cross_dict):

    grid = create_grid(size_crossword)
    read_grid(grid)
    cross_dict['grid'] = grid

    clues_across = read_across()
    clues_down = read_down()

    cross_dict['clues']['across'] = clues_across
    cross_dict['clues']['down'] = clues_down



def test_grid_coloring():
    grid = create_grid(size_crossword)
    read_grid(grid)
    print(grid)
    sys.exit()




sheet,sheet_name = ask_sheet_nr()
size_crossword = auto_determine_size()

crossword_dict = create_crossword_dict()
fill_crossword_dict(crossword_dict) 

print(f"The following crossword was read from the sheet:\n{crossword_dict}")

crossword = Crossword(crossword_dict)

'''#for debugging
while(True):
    input("press enter to solve\n")
    break
'''

solution = solve(crossword)
ask_print()




