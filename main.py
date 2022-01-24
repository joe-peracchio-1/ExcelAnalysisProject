import pandas
import tkinter as tk

window = tk.Tk()
frame_file = tk.Frame()
frame_cleaner = tk.Frame()
frame_data_subset_mode = tk.Frame()
frame_text_present_mode = tk.Frame()
frame_text_present_instances_mode = tk.Frame()
entry_file = tk.Entry(master=frame_file)
entry_cleaner_column = tk.Entry(master=frame_cleaner)
entry_cleaner_value = tk.Entry(master=frame_cleaner)
entry_subset_col = tk.Entry(master=frame_data_subset_mode)
entry_subset_continue = tk.Entry(master=frame_data_subset_mode)
entry_text_present_col = tk.Entry(master=frame_text_present_mode)
entry_text_present_value = tk.Entry(master=frame_text_present_mode)
entry_subset_value = tk.Entry(master=frame_data_subset_mode)
entry_instances_col = tk.Entry(master=frame_text_present_instances_mode)
entry_instances_value = tk.Entry(master=frame_text_present_instances_mode)
entry_cleaner_file_name = tk.Entry(master=frame_cleaner)
master_list = {}


def df_creator():
    file_name = entry_file.get()
    df = pandas.read_excel(f'{file_name}.xlsx')
    return df


def data_cleaner():
    df = df_creator()
    col_name = entry_cleaner_column.get()
    narrower = entry_cleaner_value.get()
    new_file_name = entry_cleaner_file_name.get()
    new_df = df.set_index(col_name)
    new_df = new_df.drop(narrower)
    new_df.to_excel(new_file_name)
    entry_cleaner_value.delete(0, tk.END)


def data_subset():
    data_counter_test = "Y"
    while data_counter_test == "Y":
        col_lookup = entry_subset_col.get()
        value_lookup = entry_subset_value.get()
        df = df_creator()
        df = df.set_index(col_lookup)
        df = df[df.index == value_lookup]
        master_list[value_lookup] = str(len(df.index))
        data_counter_test = entry_subset_continue.get()
        entry_subset_col.delete(0, tk.END)
        entry_subset_value.delete(0, tk.END)
        entry_subset_continue.delete(0, tk.END)
        print(master_list)
    output_file = open("outputfile.txt", "w")
    output_file.write(str(master_list))


def text_present_paragraph():
    confirmed_values = []
    df = df_creator()
    col_name = str(entry_text_present_col.get())
    check_word = str(entry_text_present_value.get())
    df = df.fillna(" ")
    text_list = df[col_name].tolist()
    for text in text_list:
        if check_word in text:
            confirmed_values.append(text)
    entry_text_present_value.delete(0, tk.END)
    entry_text_present_col.delete(0, tk.END)
    output_file = open("outputfile.txt", "w")
    for value in confirmed_values:
        output_file.write(value)
        output_file.write("\n")


def text_present_specific_instance():
    total_count = 0
    df = df_creator()
    col_name = entry_instances_col.get()
    check_word = entry_instances_value.get()
    df = df.fillna(" ")
    text_list = df[col_name].tolist()
    for text in text_list:
        if check_word in text:
            total_count += 1
    output_file = open("outputfile.txt", "w")
    output_file.write(f"Total instances of the word: {total_count}")


label_blank_1 = tk.Label()
label_blank_2 = tk.Label()
label_blank_3 = tk.Label()
label_blank_4 = tk.Label()
label_file = tk.Label(master=frame_file, text="Enter the Excel File Name")
label_cleaner = tk.Label(master=frame_cleaner, text="Data Cleaner Mode")
label_cleaner_col = tk.Label(master=frame_cleaner, text="Enter Column Name for Cleaning Mode")
label_cleaner_value = tk.Label(master=frame_cleaner, text="Enter Value to be Removed from Column")
btn_cleaner = tk.Button(master=frame_cleaner, text="Click for Cleaner Mode", command=data_cleaner)
label_subset = tk.Label(master=frame_data_subset_mode, text="Data Subset Mode")
label_subset_col = tk.Label(master=frame_data_subset_mode, text="Enter Column Name for Subset Mode")
label_subset_value = tk.Label(master=frame_data_subset_mode, text="Enter the Value to Count")
label_subset_continue = tk.Label(master=frame_data_subset_mode,
                                 text="Should the search continue. Enter 'Y' for Yes or 'N' for no")
btn_subset = tk.Button(master=frame_data_subset_mode, text="Click for Data Subset Mode", command=data_subset)
label_text_present = tk.Label(master=frame_text_present_mode, text="Text Present Mode")
label_text_present_col_name = tk.Label(master=frame_text_present_mode, text="Enter Column Name to Check for Text")
label_text_present_value = tk.Label(master=frame_text_present_mode, text="Enter Value to Check For")
btn_text_present = tk.Button(master=frame_text_present_mode, text="Click for Text Present Mode",
                             command=text_present_paragraph)
label_instances = tk.Label(master=frame_text_present_instances_mode, text="Text Instances Mode")
label_instances_col = tk.Label(master=frame_text_present_instances_mode, text="Enter Column Name to Check")
label_instances_value = tk.Label(master=frame_text_present_instances_mode, text="Enter Value to Check Number of "
                                                                                "Instances")
btn_instances = tk.Button(master=frame_text_present_instances_mode, text="Click for Text Instances Mode",
                          command=text_present_specific_instance)
label_cleaner_file_name = tk.Label(master=frame_cleaner,
                                   text="Enter Name for New Excel File (Make sure to put '.xlsx' at the end of the "
                                        "name. For example, if you want the name to be output, put down output.xlsx")
label_file.pack()
entry_file.pack()
label_subset.pack()
label_subset_col.pack()
entry_subset_col.pack()
label_subset_value.pack()
entry_subset_value.pack()
label_subset_continue.pack()
entry_subset_continue.pack()
btn_subset.pack()
label_cleaner.pack()
label_cleaner_col.pack()
entry_cleaner_column.pack()
label_cleaner_value.pack()
entry_cleaner_value.pack()
label_cleaner_file_name.pack()
entry_cleaner_file_name.pack()
btn_cleaner.pack()
label_text_present.pack()
label_text_present_col_name.pack()
entry_text_present_col.pack()
label_text_present_value.pack()
entry_text_present_value.pack()
btn_text_present.pack()
label_instances.pack()
label_instances_col.pack()
entry_instances_col.pack()
label_instances_value.pack()
entry_instances_value.pack()
btn_instances.pack()
file_name = entry_file.get()
frame_file.pack()
label_blank_1.pack()
frame_data_subset_mode.pack()
label_blank_2.pack()
frame_cleaner.pack()
label_blank_3.pack()
frame_text_present_mode.pack()
label_blank_4.pack()
frame_text_present_instances_mode.pack()
window.mainloop()
