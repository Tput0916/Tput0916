import os
import re
import cv2
import pytesseract
from pytesseract import Output
from fuzzywuzzy import fuzz
import pandas as pd
from itertools import product

def extract_value(image, params):
    """A generic extraction function that uses a given set of parameters."""
    try:
        # --- 1. Preprocessing with dynamic parameters ---
        h, w = image.shape[:2]
        y_start, y_end = params['crop_y']
        x_start, x_end = 0, w
        
        cropped = image[y_start:y_end, x_start:x_end]
        gray = cv2.cvtColor(cropped, cv2.COLOR_BGR2GRAY)
        
        scaled = cv2.resize(gray, None, fx=params['scale'], fy=params['scale'], interpolation=cv2.INTER_CUBIC)
        
        if params['blur'] > 0:
            scaled = cv2.GaussianBlur(scaled, (params['blur'], params['blur']), 0)

        # --- 2. OCR with dynamic PSM ---
        config = f"--psm {params['psm']}"
        data = pytesseract.image_to_data(scaled, lang='eng', config=config, output_type=Output.DICT)
        
        # --- 3. Structural Analysis Logic (as before) ---
        lines = {}
        for i in range(len(data['text'])):
            if int(data['conf'][i]) > 40:
                line_num = data['line_num'][i]
                if line_num not in lines: lines[line_num] = []
                lines[line_num].append({'text': data['text'][i].strip(), 'left': data['left'][i], 'top': data['top'][i], 'width': data['width'][i], 'height': data['height'][i]})

        header_row, data_row = None, None
        for _, words in lines.items():
            line_text = " ".join([w['text'] for w in words])
            if fuzz.partial_ratio("Average Minimum Maximum", line_text) > 80: header_row = words
            if fuzz.partial_ratio("All Pairs", line_text) > 80: data_row = words
        
        if not header_row or not data_row: return None

        avg_word = max(header_row, key=lambda w: fuzz.ratio("Average", w['text']))
        pairs_word = data_row[0]

        col_x_start, col_x_end = avg_word['left'] - 20, avg_word['left'] + avg_word['width'] + 20
        row_y_start, row_y_end = pairs_word['top'] - 20, pairs_word['top'] + pairs_word['height'] + 20

        parts = sorted([w for _, words in lines.items() for w in words if col_x_start <= (w['left'] + w['width']/2) <= col_x_end and row_y_start <= (w['top'] + w['height']/2) <= row_y_end and re.search(r'[\d.]', w['text'])], key=lambda p: p['left'])
        
        if not parts: return None
        
        num_str = re.sub(r'[^\d.]', '', "".join([p['text'] for p in parts]))
        if num_str and '.' in num_str: return float(num_str)

    except Exception:
        return None
    return None

def main():
    """Finds the best OCR parameters by testing against a ground truth file."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    # --- Load Ground Truth Data ---
    try:
        truth_df = pd.read_csv('ground_truth.csv', header=None, names=['filename', 'truth_value'])
        truth_data = {row.filename: row.truth_value for _, row in truth_df.iterrows()}
    except FileNotFoundError:
        print("Error: ground_truth.csv not found.")
        return

    # --- Define Parameter Space ---
    param_space = {
        'crop_y': [(100, 400), (150, 350)],
        'scale': [2.0, 2.5, 3.0],
        'blur': [0, 3], # 0 means no blur
        'psm': [3, 6, 11]
    }
    
    param_combinations = [dict(zip(param_space.keys(), v)) for v in product(*param_space.values())]
    
    best_params = None
    max_accuracy = -1

    print(f"Testing {len(param_combinations)} parameter combinations against {len(truth_data)} images...")

    # --- Iterate and Find Best Parameters ---
    for i, params in enumerate(param_combinations):
        correct_count = 0
        for filename, true_value in truth_data.items():
            image = cv2.imread(filename)
            if image is None: continue
            
            extracted_value = extract_value(image, params)
            
            if extracted_value is not None:
                # Check if the extracted value is close to the true value
                if abs(extracted_value - true_value) < 0.01:
                    correct_count += 1
        
        accuracy = correct_count / len(truth_data)
        print(f"Combination {i+1}/{len(param_combinations)} -> Accuracy: {accuracy:.2%}")

        if accuracy > max_accuracy:
            max_accuracy = accuracy
            best_params = params
            print(f"  *** New best parameters found: {best_params} with accuracy {max_accuracy:.2%} ***")

    print("\n--- Optimization Complete ---")
    print(f"Best Parameters Found: {best_params}")
    print(f"Highest Accuracy on Test Set: {max_accuracy:.2%}")

if __name__ == '__main__':
    main()