import pandas as pd
from collections import OrderedDict, defaultdict

file_path = 'All_Inventory - Keystone Medium.xlsx'
sheet_name = "Online PGM"
start_col, end_col = 1, 61

# sheet_names = ['CNTL Card', 'Online PGM', 'Batch PGM']
# start_end_col = [[14, 88], [21, 90], [14, 67]]
# file_path = 'All_Inventory - Keystone WITHOUT ZEROSTAT.xlsx'
# sheet_name = 'Online PGM'  # can change sheet name
# start_col, end_col = 21, 90 # start_col = column number for first file end_col = last file column number


print(f"File Name: {file_path}\nSheet Name: {sheet_name}\nStart-End Column: {start_col, end_col}\n")
dec = "\n**************************************************************************************\n"

df = pd.read_excel(file_path, sheet_name=sheet_name)
df.columns = df.columns.map(str).str.strip()

component_col = [col for col in df.columns if 'Component' in col and 'Name' in col][0]
df.set_index(component_col, inplace=True)

df = df[~df.index.to_series().fillna("").str.lower().str.contains("total")]
df = df[~df.index.isna()]
df = df.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

selected_columns = df.columns[start_col:end_col+1]
df_files = df[selected_columns]

print(f"Files {len(selected_columns)}:\n {selected_columns}")
print(f"\nPrograms {len(df_files)}:\n {df_files.index}")
print(dec)

completely_unused_files = df_files.columns[df_files.sum(axis=0) == 0].tolist()
programs_with_no_files = df_files.index[df_files.sum(axis=1) == 0].tolist()

df_files = df_files.drop(columns=completely_unused_files)
df_files = df_files.drop(index=programs_with_no_files)

program_to_files = {}
for program, row in df_files.iterrows():
    used_files = list(row[row == 1].index)
    program_to_files[program] = used_files

sorted_programs = sorted(program_to_files.items(), key=lambda x: len(x[1]), reverse=True)
remaining_programs = OrderedDict(sorted_programs)

# OUTPUT: Program Name: [list of files used by that program]
print("\n=== Programs and Files Used (Sorted by File Count Descending) ===\n")
for prog, files in sorted_programs:
    print(f"{prog}: [{', '.join(sorted(files))}]")

covered_files = set()
group_list = []

while remaining_programs:
    valid_program = None
    for prog, files in remaining_programs.items():
        new_files = [f for f in files if f not in covered_files]
        if new_files:
            valid_program = prog
            break
    if not valid_program:
        break

    group_files = set(f for f in remaining_programs[valid_program] if f not in covered_files)

    # Split into chunks if more than 25 files
    file_chunks = []
    group_files_list = list(group_files)
    while group_files_list:
        chunk = group_files_list[:25]
        file_chunks.append(set(chunk))
        group_files_list = group_files_list[25:]

    for chunk in file_chunks:
        cleared_programs = []
        for prog, files in remaining_programs.items():
            if set(files).issubset(covered_files.union(chunk)) and set(files).issubset(chunk.union(covered_files)):
                cleared_programs.append(prog)

        group_list.append({
            "files": chunk,
            "programs": cleared_programs
        })

        covered_files.update(chunk)
        for prog in cleared_programs:
            remaining_programs.pop(prog)

# Merging small groups into larger ones based on overlap
merged_group_list = []
for group in group_list:
    if len(group['files']) > 1:
        merged_group_list.append(group)

# Create a mapping from file to group index for larger groups
file_to_group = defaultdict(list)
for idx, group in enumerate(merged_group_list):
    for file in group['files']:
        file_to_group[file].append(idx)

# Now check small groups (len==1) and try to merge
for group in group_list:
    if len(group['files']) == 1:
        file = list(group['files'])[0]
        progs = group['programs']
        best_group_idx = None
        best_overlap = 0
        for idx, big_group in enumerate(merged_group_list):
            overlap = 0
            for prog in progs:
                if prog in program_to_files:
                    used = set(program_to_files[prog])
                    overlap += len(used & big_group['files'])
            if overlap > best_overlap:
                best_overlap = overlap
                best_group_idx = idx
        if best_group_idx is not None:
            merged_group_list[best_group_idx]['files'].add(file)
            merged_group_list[best_group_idx]['programs'].extend(progs)
        else:
            merged_group_list.append(group)

# Final cleanup to remove duplicate programs in merged groups
for group in merged_group_list:
    group['programs'] = list(set(group['programs']))

# Sort groups by file count descending
merged_group_list.sort(key=lambda g: len(g['files']), reverse=True)

# OUTPUT: Group with file usage
print(dec)
for i, group in enumerate(merged_group_list, start=1):
    files = sorted(group['files'])
    programs = group['programs']
    print(f"Group {i} (with {len(files)} files): [{', '.join(files)}]")
    print(f"Used by {len(programs)} programs:")
    
    sorted_programs_in_group = sorted(programs, key=lambda p: len(program_to_files.get(p, [])), reverse=True)
    for prog in sorted_programs_in_group:
        used_files = program_to_files.get(prog, [])
        
        # BELOW LINE WILL PRINT ONLY PROGRAM NAME; IF WANT TO PRINT FILE USED BY IT COMMENT below one AND UNCOMMENT the next one
        print(f" - {prog}")
        # print(f" - {prog}: [{', '.join(sorted(used_files))}]")

    print(dec)

if completely_unused_files:
    print(f"Files not used by any program : {len(completely_unused_files)}")
    for f in sorted(completely_unused_files):
        print(f" - {f}")
else:
    print("\nFiles not used by any program : 0")

print(dec)
if programs_with_no_files:
    print(f"Programs with no file usage : {len(programs_with_no_files)}")
    for prog in sorted(programs_with_no_files):
        print(f" - {prog}")
else:
    print("Programs with no file usage : 0")


# Summary
total_programs_in_dataset = len(program_to_files) + len(programs_with_no_files)
total_programs_in_groups = sum(len(group['programs']) for group in merged_group_list)
total_programs_no_files = len(programs_with_no_files)

all_files_used_in_groups = set()
for group in merged_group_list:
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


# Validation
assert total_programs_in_dataset == total_programs_in_groups + total_programs_no_files, \
    "Mismatch in program count! Check for duplication or missing programs."
assert total_files_in_dataset == total_files_in_groups + total_files_never_used, \
    "Mismatch in file count! Check for duplication or unused files."
    
print("\nAssertion Check Passed: All programs and files are fully accounted.\n")
