import originpro as op
import PyOrigin
import pandas as pd
import numpy as np
from scipy import stats

#==================================
# 1. CONFIGURATION & INITIALIZATION
#==================================
note = op.new_notes('Results')
results_dict = {}
bd = float(input("Please enter the initial potential (V): "))
kt = float(input("Please enter the final potential (V): "))

#==================================
# 2. CONFIGURATION & INITIALIZATION
#==================================
def count_open_workbooks():
    open_count = 0
    book_index = 1
    while True:
        book_name = f"Book{book_index}"
        wb = op.find_book('w', book_name)
        if wb:
            if wb.is_open():
                open_count += 1
            book_index += 1
        else:
            break
    print(f"Total active workbooks: {open_count}")
    note.append(f"Total active workbooks: {open_count}")
    note.append(f"Chosen potential range to calculate linear line: {bd} to {kt} (V)")
    note.append("------------------------------")
    return open_count
def process_workbook(book_name):
    print(f"\nProcessing {book_name}...")
    note.append(f"\nProcessing {book_name}...")
    wb = op.find_book('w', book_name)
    wb.activate()
    source_wks = wb[0]
    source_wks.activate()
    
    #find peak
    df = source_wks.to_df()
    df_selected = df[["Vf", "Im"]]
    peak_index = df_selected["Im"].idxmax()
    peak_vf = df_selected["Vf"].iloc[peak_index]
    peak_im = df_selected["Im"].iloc[peak_index]
    
    #draw graph
    source_wks.cols_axis('X', 'D')
    source_wks.cols_axis('Y', 'E')
    graph = op.new_graph(template='line')
    gl = graph[0]
    _thongso = f'{source_wks.lt_range()}!(?,4:5)'
    plot = gl.add_plot(_thongso)
    lgnd = gl.label('Legend')
    lgnd.set_int('fsize', 22)
    lgnd.set_int('left', 1400)
    lgnd.set_int('top', 700)
    lgnd.set_int('showframe', 0)
    graph.activate()
    Vf_min_index = df_selected["Vf"].idxmin()
    min_Vf = df_selected["Vf"].iloc[Vf_min_index]
    canhtrai = round (min_Vf - 0.05, 2)
    Vf_max_index = df_selected["Vf"].idxmax()
    max_Vf = df_selected["Vf"].iloc[Vf_max_index]
    canhphai = round (max_Vf + 0.05, 2)
    graph[0].set_xlim(canhtrai, canhphai, 0.1)
    canhtren = peak_im + 0.00003
    kcy = round(((canhtren - (-0.00003) ) / 4), 5)
    graph[0].set_ylim(-0.00003, canhtren, kcy)
    lgnd.text='\l(1) LSV curve'
    
    #linear
    linear_region = df_selected[(df_selected["Vf"] >= bd) & (df_selected["Vf"] <= kt)]
    slope, intercept, _, _, _ = stats.linregress(linear_region["Vf"], linear_region["Im"])
    note.append(f"Linear line: Y = {slope}X + {intercept};")
    
    #draw linear
    y1 = slope * bd + intercept
    y2 = slope * 0.3 + intercept
    _dfvett = pd.DataFrame({
    'X': [bd,0.3],
    'Linear':[y1,y2],
    })
    new_wks = wb.add_sheet('DataDrawLine')
    new_wks.from_df(_dfvett)
    _thongso2 = f'{new_wks.lt_range()}!(?,1:end)'
    plot2 = gl.add_plot(_thongso2)
    plot2.color = '#f00'
    source_wks.activate()
    
    #peak
    note.append(f'Peak (Vf, Im): {peak_vf} V, {peak_im} A;')
    
    #Find distance
    y_on_line = slope * peak_vf + intercept
    distance = abs(peak_im - y_on_line)
    results_dict[book_name] = distance
    print(f"Distance from peak to linear line: {distance} A")
    note.append(f"Distance from peak to linear line: {distance} A;")
    
    #draw distance
    y3 = slope * peak_vf + intercept
    _dfvekc = pd.DataFrame({
    'X': [peak_vf,peak_vf],
    'Distance':[peak_im,y3],
    })
    new_wks2 = wb.add_sheet('DataDistance')
    new_wks2.from_df(_dfvekc)
    _thongso3 = f'{new_wks2.lt_range()}!(?,1:end)'
    plot3 = gl.add_plot(_thongso3)
    plot3.color = '#00f'
    source_wks.activate()

#====================
# 3. EXERCUTION BLOCK
#====================
def process_all_open_workbooks():
    open_count = count_open_workbooks()
    if open_count > 0:
        book_index = 1
        while True:
            book_name = f"Book{book_index}"
            wb = op.find_book('w', book_name)
            if wb and wb.is_open():
                process_workbook(book_name)
            else:
                break
            book_index += 1
process_all_open_workbooks()

#===============
# 4. MEAN AND SD
#===============
if results_dict:
    values = list(results_dict.values())
    mean_value = np.mean(values)
    std_dev = np.std(values, ddof=1)
    mean_value2 = round(mean_value, 8)
    mean_value3 = round(mean_value*1000000,2)
    print(f"\nMean of all peaks: {mean_value2} A = {mean_value3} uA")
    print(f"Standard deviation of all peaks: {std_dev}")
    note.append(f"\n--------------------------------------------")
    note.append(f"Mean of all peaks: {mean_value2} A = {mean_value3} uA")
    note.append(f"Standard deviation of all peaks: {std_dev}")
    note.append(f"--------------------------------------------")
else:
    print("\nNo valid data to calculate statistics.")