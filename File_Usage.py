import pandas as pd
from collections import OrderedDict

file_path = 'All_Inventory - Keystone Medium.xlsx'
sheet_name = "Online PGM"
start_col, end_col = 1, 61

# sheet_names = ['CNTL Card', 'Online PGM', 'Batch PGM']
# start_end_col = [[14, 88], [21, 90], [14, 67]]
# file_path = 'All_Inventory - Keystone WITHOUT ZEROSTAT.xlsx'
# sheet_name = 'Online PGM'  # can change sheet name
# start_col, end_col = 21, 90 # start_col = column number for first file end_col = last file column number

dec = "\n**************************************************************************************\n"

df = pd.read_excel(file_path, sheet_name=sheet_name)
df.columns = df.columns.map(str).str.strip()

# Identify Component Name (programs)
component_col = [col for col in df.columns if 'Component' in col and 'Name' in col][0]
df.set_index(component_col, inplace=True)

# Clean data
df = df[~df.index.to_series().fillna("").str.lower().str.contains("total")]
df = df[~df.index.isna()]
df = df.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)


selected_columns = df.columns[start_col:end_col+1]
df_files = df[selected_columns]

# files and programs with no usage
completely_unused_files = df_files.columns[df_files.sum(axis=0) == 0].tolist()
programs_with_no_files = df_files.index[df_files.sum(axis=1) == 0].tolist()

# Remove them before processing 
df_files = df_files.drop(columns=completely_unused_files)
df_files = df_files.drop(index=programs_with_no_files)

# Map programs → used files
program_to_files = {}
for program, row in df_files.iterrows():
    used_files = list(row[row == 1].index)
    program_to_files[program] = used_files

# Sort programs by number of files used to take first group with most files
sorted_programs = sorted(program_to_files.items(), key=lambda x: len(x[1]), reverse=True)
remaining_programs = OrderedDict(sorted_programs)

# Print all programs and their used files
print("\n=== Programs and Files Used (Sorted by File Count Descending) ===\n")
for prog, files in sorted_programs:
    print(f"{prog}: [{', '.join(sorted(files))}]")

# Initialize for grouping
covered_files = set()
group_list = []

# Grouping
while remaining_programs:
    # Find first program with uncovered files
    valid_program = None
    for prog, files in remaining_programs.items():
        new_files = [f for f in files if f not in covered_files]
        if new_files:
            valid_program = prog
            break
    if not valid_program:
        break

    group_files = set(f for f in remaining_programs[valid_program] if f not in covered_files)

    # Find all programs using only these files (direct subset)
    cleared_programs = []
    for prog, files in remaining_programs.items():
        if set(files).issubset(covered_files.union(group_files)):
            cleared_programs.append(prog)

    # Save group
    group_list.append({
        "files": group_files,
        "programs": cleared_programs
    })

    # Update covered files and remove programs
    covered_files.update(group_files)
    for prog in cleared_programs:
        remaining_programs.pop(prog)

# Sort groups by file count descending
group_list.sort(key=lambda g: len(g['files']), reverse=True)

# Print grouped results 
print(dec)
# print("Final File Groups (Sorted by File Count Descending)\n")
for i, group in enumerate(group_list, start=1):
    files = sorted(group['files'])
    programs = group['programs']
    print(f"Group {i} (with {len(files)} files): [{', '.join(files)}]")
    print(f"Used by {len(programs)} programs:")
    for prog in programs:
        used_files = program_to_files.get(prog, [])
        print(f" - {prog}: [{', '.join(sorted(used_files))}]")

    print(dec)

# Print unused files (that were removed)
if completely_unused_files:
    print(f"Files not used by any program : {len(completely_unused_files)}")
    for f in sorted(completely_unused_files):
        print(f" - {f}")
else:
    print("\nFiles not used by any program : 0")

# print programs that use no files (that were removed) 
print(dec)
if programs_with_no_files:
    print(f"Programs with no file usage : {len(programs_with_no_files)}")
    for prog in sorted(programs_with_no_files):
        print(f" - {prog}")
else:
    print("Programs with no file usage : 0")


# Summary

total_programs_in_dataset = len(program_to_files) + len(programs_with_no_files)
total_programs_in_groups = sum(len(group['programs']) for group in group_list)
total_programs_no_files = len(programs_with_no_files)

all_files_used_in_groups = set()
for group in group_list:
    all_files_used_in_groups.update(group['files'])

total_files_in_dataset = len(df_files.columns) + len(completely_unused_files)
total_files_in_groups = len(all_files_used_in_groups)
total_files_never_used = len(completely_unused_files)

print(dec)
print(f" Total Programs: {total_programs_in_dataset}")
print(f"  - In groups  : {total_programs_in_groups}")
print(f"  - With no use: {total_programs_no_files}")

print(dec)
print(f" Total Files  : {total_files_in_dataset}")
print(f"  - In groups  : {total_files_in_groups}")
print(f"  - Never used : {total_files_never_used}")

# Validations 

assert total_programs_in_dataset == total_programs_in_groups + total_programs_no_files, \
    "Mismatch in program count! Check for duplication or missing programs."
assert total_files_in_dataset == total_files_in_groups + total_files_never_used, \
    "Mismatch in file count! Check for duplication or unused files."
    
print("\nAssertion Check Passed: All programs and files are fully accounted.\n")