import pandas as pd
import numpy as np

# === Load Data ===
file_path = 'All_Inventory - Keystone Medium.xlsx'
sheet_name = "Online PGM"
start_col, end_col = 1, 61

# sheet_names = ['CNTL Card', 'Online PGM', 'Batch PGM']
# start_end_col = [[14, 88], [21, 90], [14, 67]]
# file_path = 'All_Inventory - Keystone WITHOUT ZEROSTAT.xlsx'
# sheet_name = 'Online PGM'  # can change sheet name
# start_col, end_col = 21, 90 # start_col = column number for first file end_col = last file column number

print(f"File Name: {file_path}\nSheet Name: {sheet_name}\nStart-End Column: {start_col, end_col}\n")

# === Read Excel ===
df = pd.read_excel(file_path, sheet_name=sheet_name)
df.columns = df.columns.map(str).str.strip()

# Identify the column that contains program names
component_col = [col for col in df.columns if 'Component' in col and 'Name' in col][0]
df.set_index(component_col, inplace=True)

# Clean the data
df = df[~df.index.to_series().fillna("").str.lower().str.contains("total")]
df = df[~df.index.isna()]
df = df.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

# Select only file columns (1s and 0s indicating usage)
selected_columns = df.columns[start_col:end_col+1]
df_files = df[selected_columns]

print(f"Files {len(selected_columns)}:\n {selected_columns}")
print(f"\nPrograms {len(df_files)}:\n {df_files.index}")
print("\n" + "-" * 100 + "\n")

# === 1. Pairwise Commonality Matrix ===
commonality_matrix = df_files.T.dot(df_files)
np.fill_diagonal(commonality_matrix.values, 0)

# === 2. Files grouped by total commonality score ===
commonality_scores = commonality_matrix.sum(axis=1)
grouped_files = commonality_scores.sort_values(ascending=False)
sorted_files = grouped_files.index
commonality_matrix_sorted = commonality_matrix.loc[sorted_files, sorted_files]

print("\n=== 1. Sorted Pairwise File Commonality Matrix ===\n")
print(commonality_matrix_sorted)
print("\n" + "-" * 100 + "\n")


print("\n=== 2. Files Grouped by Total Commonality Score (Descending) ===\n")
print(grouped_files)
print("\n" + "-" * 100 + "\n")


# === 3. Programs and Files Used (Sorted by File Count Descending) ===
print("\n=== 3. Programs and Files Used (Sorted by File Count Descending) ===\n")
program_file_map = {}
programs_with_no_use = []


for program, row in df_files.iterrows():
    files_used = sorted(list(row[row == 1].index))
    if files_used:
        program_file_map[program] = files_used
    else:
        programs_with_no_use.append(program)

program_file_map_sorted = dict(sorted(program_file_map.items(), key=lambda item: len(item[1]), reverse=True))

for prog, files in program_file_map_sorted.items():
    print(f"{prog} - {files}")
print("\n" + "-" * 100 + "\n")

# === 4. File Groups and Matching Programs ===
print("\n=== 4. File Co-Usage Grouping ===\n")
counter = 1
program_count = 0
printed_programs = set()
all_seen_files = set()
largest_group = []
largest_group_len = 0

for file in grouped_files.index:
    related_files = commonality_matrix.loc[file]
    used_with = [f for f in related_files.index if related_files[f] > 0]
    used_with = [file] + used_with
    used_with_set = set(used_with)

    matching_programs = []
    for program, row in df_files.iterrows():
        if program in printed_programs:
            continue
        files_used = set(row[row == 1].index)
        if files_used and files_used.issubset(used_with_set):
            matching_programs.append(program)
            printed_programs.add(program)

    if matching_programs:
        new_files = sorted(used_with_set - all_seen_files)
        all_seen_files.update(used_with_set)

        print(f"Group {counter}: {len(used_with)} Files")
        
        if counter!=1:
            print(f"\nNewly added files = {new_files}\n")
        
        # print(sorted(used_with))
        print(used_with)
        print()
        
        print(f"{len(matching_programs)} Programs using ONLY the above {len(used_with)} files:")
        print(matching_programs)
        print("\n" + "-" * 100 + "\n")

        counter += 1
        program_count += len(matching_programs)

        if len(used_with) > largest_group_len:
            largest_group_len = len(used_with)
            largest_group = used_with


# === 5. Disjoint File Groups ===
print("\n=== 5. Disjoint File Groups (if any) ===\n")
largest_group_set = set(largest_group)
disjoint_groups = []
remaining_files = set(commonality_matrix.index) - largest_group_set
visited = set()

for file in remaining_files:
    if file in visited:
        continue
    related = commonality_matrix.loc[file]
    connected = set([f for f in related.index if related[f] > 0 and f in remaining_files])
    connected.add(file)
    visited |= connected

    if connected and connected.isdisjoint(largest_group_set) and len(connected) > 1:
        disjoint_groups.append(connected)

if disjoint_groups:
    print(f"Found {len(disjoint_groups)} disjoint group(s) of files:\n")
    for i, group in enumerate(disjoint_groups, 1):
        print(f"Group {i} ({len(group)} files): {sorted(group)}\n")
else:
    print("No completely disjoint file group found other than the largest group.\n")

# === 6. Summary and Assertions ===
print("\n" + "*" * 90 + "\n")

# Files never used
files_never_used = [col for col in df_files.columns if df_files[col].sum() == 0]
print(f"Files not used by any program: {len(files_never_used)}")
if files_never_used:
    print(" - " + ", ".join(files_never_used))

print("\n" + "*" * 90 + "\n")

# Programs with no file usage
print(f"Programs with no file usage : {len(programs_with_no_use)}")
for p in programs_with_no_use:
    print(f" - {p}")

print("\n" + "*" * 90 + "\n")

# Final summary
total_programs = len(df_files)
programs_in_groups = len(printed_programs)
programs_with_no_use_count = len(programs_with_no_use)

total_files = len(df_files.columns)
files_in_groups = total_files - len(files_never_used)

print(" Total Programs:", total_programs)
print("  - In groups  :", programs_in_groups)
print("  - With no use:", programs_with_no_use_count)

print("\n" + "*" * 90 + "\n")

print(" Total Files  :", total_files)
print("  - In groups  :", files_in_groups)
print("  - Never used :", len(files_never_used))

# === 7. Final Assertions ===
assert programs_in_groups + programs_with_no_use_count == total_programs, "Program count mismatch!"
assert files_in_groups + len(files_never_used) == total_files, "File count mismatch!"
print("\nAssertion Check Passed: All programs and files are fully accounted.\n")
