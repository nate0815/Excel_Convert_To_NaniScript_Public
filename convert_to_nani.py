import os
import pandas as pd
import re

FINAL_COMMAND = "@stopAvg"

# Switch to the current .py path
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
print(f"📁 腳本執行路徑：{script_dir}")

# Find Excel file
excel_file = next((f for f in os.listdir('.') if f.endswith('.xlsx')), None)
if not excel_file:
    print("❌ 找不到 Excel 檔案 (.xlsx)")
    exit()
print(f"📄 已找到 Excel 檔案：{excel_file}")

# Load Excel
xls = pd.ExcelFile(excel_file)
sheet_names = xls.sheet_names

# Load character mapping table, increase robustness
character_map = {}
character_portrait_info = {} # New map: char_id -> has_portrait (bool)
if 'Character' in sheet_names:
    try:
        character_df = xls.parse('Character')
        # Ensure columns exist, remove rows containing NaN, and convert to string to prevent issues
        if '中文顯示' in character_df.columns and 'id' in character_df.columns:
            character_df = character_df.dropna(subset=['中文顯示', 'id'])
            character_df['中文顯示'] = character_df['中文顯示'].astype(str).str.strip()
            character_df['id'] = character_df['id'].astype(str).str.strip()

            # Handle '是否有立繪' column for portrait display logic
            has_portrait_col = '是否有立繪'
            if has_portrait_col in character_df.columns:
                # Convert to string, fill NaN with 'T' (default to has portrait), and normalize
                character_df[has_portrait_col] = character_df[has_portrait_col].fillna('T').astype(str).str.strip().str.upper()
            else:
                # If column doesn't exist, assume all characters have portraits for backward compatibility
                character_df[has_portrait_col] = 'T'

            character_map = dict(zip(character_df['中文顯示'], character_df['id']))
            character_portrait_info = dict(zip(character_df['id'], character_df[has_portrait_col] != 'F'))

            if not character_map:
                 print("⚠️ 'Character' 分頁已載入，但未成功建立任何角色對應（可能為空或格式問題）。")
        else:
            print("⚠️ 'Character' sheet is missing '中文顯示' or 'id' columns.")
    except Exception as e:
        print(f"⚠️ Cannot parse 'Character' sheet or create character mapping: {e}")
else:
    print("⚠️ 找不到 'Character' 分頁，將無法對應角色 ID。")

# Prepare output folder
output_folder = os.path.splitext(excel_file)[0]
os.makedirs(output_folder, exist_ok=True)

all_generated_lines = {} # Store initially processed command lines for each sheet

for sheet in sheet_names:
    if sheet in ['Character', 'Stage']:
        continue

    print(f"\n🔧 處理分頁：{sheet}")
    try:
        # Use fillna('') to convert all NaN values to empty strings, avoiding str(NaN) becoming "nan" later
        df = xls.parse(sheet).fillna('')
    except Exception as e:
        print(f"⚠️ 無法解析 {sheet}：{e}")
        continue

    # Check for necessary columns, '對話內容' (Dialogue Content) is core
    if '對話內容' not in df.columns:
        print(f"⚠️ 分頁 {sheet} 缺少核心的 '對話內容' 欄位，跳過此分頁。")
        continue
    # '角色' (Character) and '選項' (Option) columns are optional; if they don't exist, .get('', '') will handle them

    lines = []
    prev_char_id = None
    currently_in_choice_block = False # True if processing a sequence of choices
    processed_any_choice_in_sheet = False # True if any choice was made in this sheet

    for index, row in df.iterrows():
        # .get(column, '') returns '' if the column doesn't exist or its value is NaN (which has already been converted to '' by fillna(''))
        # str() ensures it's a string, .strip() removes leading/trailing spaces
        speaker_raw = row.get('角色', '') # fillna('') has handled NaN
        speaker = str(speaker_raw).strip()

        text = str(row.get('對話內容', '')).strip()

        option_goto = ''
        original_option_cell_value = '' # Used to determine if the '選項' (Option) column was originally empty or just spaces
        if '選項' in df.columns: # Check if the '選項' (Option) column exists
            original_option_cell_value = str(row.get('選項', '')) # Already a string due to fillna
            option_goto = original_option_cell_value.strip()

        # 1. Process options: '對話內容'(text) as option text, '選項'(option_goto) as jump target
        if text and option_goto:
            if not currently_in_choice_block: # First choice in a new block
                if prev_char_id: # If a character was speaking before this choice block
                    if character_portrait_info.get(prev_char_id, True): # Check if character has a portrait
                        lines.append(f'@hide {prev_char_id}')
                currently_in_choice_block = True
            processed_any_choice_in_sheet = True
            lines.append(f'@choice "{text}" goto:.{option_goto}') # New choice format
            prev_char_id = None # Choices reset speaker context
            continue

        # If we were in a choice block, and this line is NOT a choice, the block ends.
        if currently_in_choice_block: # (and this line is not a choice, because 'continue' was hit if it was)
            lines.append('@stop') # Add @stop after the choice block
            currently_in_choice_block = False
            # prev_char_id is already None here.

        # 2. Handle warnings for invalid options:
        # (This logic remains, but it's now after the choice block termination)
        if text and not option_goto and '選項' in df.columns and original_option_cell_value != '':
            print(f"⚠️ 第 {index + 2} 行 (分頁 {sheet})：'對話內容' ('{text}') 存在，但 '選項' 欄 (原始值: '{original_option_cell_value}') 內容無效或轉換後為空。此行已跳過。") # User message, kept in Chinese
            continue

        # 3. If '對話內容'(text) is empty, skip this line
        if not text:
            continue

        # At this point, 'text' is guaranteed to be non-empty, 
        # and this line is not a valid option, not a skipped invalid option line, 
        # and not part of a choice block that just ended.

        # 4. Process narration: if speaker is empty (and text is non-empty)
        if not speaker:
            if prev_char_id: # If the previous line was a character speaking, hide them first
                if character_portrait_info.get(prev_char_id, True): # Check if character has a portrait
                    lines.append(f'@hide {prev_char_id}')
            lines.append(text)
            prev_char_id = None # After narration, there is no current character
            continue

        # 5. Process character speech: (text non-empty, speaker non-empty)
        char_id = character_map.get(speaker)
        if not char_id:
            print(f"⚠️ 第 {index + 2} 行 (分頁 {sheet})：找不到角色 ID '{speaker}' 對應的 Naninovel ID (對話: '{text}')。請檢查 'Character' 分頁。") # User message
            continue # Skip this line with an unrecognized character speaking

        if prev_char_id and prev_char_id != char_id:
            if character_portrait_info.get(prev_char_id, True): # Check if previous character has a portrait
                lines.append(f'@hide {prev_char_id}')
        
        # Show @char only if the current character is different from the previous one, 
        # or if there was no previous character.
        # Also handles the first appearance of a character ID.
        if prev_char_id != char_id:
            if character_portrait_info.get(char_id, True): # Check if current character has a portrait
                lines.append(f'@char {char_id}')

        lines.append(f'{char_id}: {text}')
        prev_char_id = char_id

    # After the loop finishes
    if currently_in_choice_block: # If the sheet ended with choices
        lines.append('@stop')

    if prev_char_id: # If the last line was character speech, add @hide
        if character_portrait_info.get(prev_char_id, True): # Check if character has a portrait
            lines.append(f'@hide {prev_char_id}')

    # Add FINAL_COMMAND if the sheet has content AND it did NOT process any choices.
    if lines and not processed_any_choice_in_sheet:
        lines.append(FINAL_COMMAND)

    all_generated_lines[sheet] = lines

# --- Phase 2: Iteratively merge scripts with @choice and @goto ---
scripts_data = {name: list(lines_list) for name, lines_list in all_generated_lines.items()} # Operate on a copy
final_merged_sheets = set() # Record sheet names that have been merged into other scripts

# Regex to parse @choice "text" goto:.Target
choice_goto_pattern = re.compile(r'@choice\s+"[^"]*"\s+goto:\.(.+)')

while True:
    merges_in_this_iteration = 0
    
    # Process in a fixed order (sorted sheet names) to ensure consistent output, 
    # although the order has minimal impact here.
    sorted_sheet_names = sorted(list(scripts_data.keys()))

    for source_name in sorted_sheet_names:
        if source_name in final_merged_sheets: # If this sheet itself has already been merged, skip it
            continue

        current_lines_for_source = list(scripts_data[source_name]) # Use a copy
        
        # Lines that constitute the source script itself, after processing its choices for potential merges.
        source_script_intrinsic_lines = [] 
        # Content blocks of scripts merged INTO this source. Each item is a list of lines.
        content_of_merged_targets = []
        
        made_a_merge_in_this_source = False

        for line in current_lines_for_source:
            match = choice_goto_pattern.match(line)
            if match:
                target_name = match.group(1).strip() # group(1) is the target name after goto:.

                # Check if this target can and should be merged
                if target_name in scripts_data and \
                   target_name != source_name and \
                   target_name not in final_merged_sheets:

                    source_script_intrinsic_lines.append(line) # Keep the @choice line

                    # Prepare the block for this target
                    target_actual_content = scripts_data[target_name]
                    merged_block = []
                    merged_block.append(f'; User marker: [# {target_name}]')
                    merged_block.append(f'# {target_name}')
                    merged_block.extend(target_actual_content)
                    content_of_merged_targets.append(merged_block)

                    final_merged_sheets.add(target_name) # Mark as merged
                    merges_in_this_iteration += 1
                    made_a_merge_in_this_source = True
                else:
                    # Choice line, but target not mergeable (e.g. already merged, non-existent, self-reference)
                    # Keep the choice line as is.
                    source_script_intrinsic_lines.append(line)
            else:
                # Not a @choice goto line, keep it
                source_script_intrinsic_lines.append(line)
        
        if made_a_merge_in_this_source:
            final_lines_for_source = list(source_script_intrinsic_lines) 
            for block in content_of_merged_targets:
                final_lines_for_source.extend(block)
            scripts_data[source_name] = final_lines_for_source

    if merges_in_this_iteration == 0: # If no new sheets were merged in this entire iteration, then end
        break

# --- Phase 3: Output final .nani files ---
for sheet_name, lines_content in scripts_data.items():
    if sheet_name not in final_merged_sheets: # Only output scripts that were not merged into other files
        output_path = os.path.join(output_folder, f"{sheet_name}.nani")
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines_content))
        print(f"✅ 匯出 ({sheet_name})：{output_path}")

# Automatically open the output folder (Windows)
try:
    os.startfile(output_folder)
except AttributeError: # os.startfile is Windows only
    print(f"\nℹ️ 自動開啟資料夾功能僅支援 Windows。")
    print(f"📁 請手動開啟：{os.path.abspath(output_folder)}")
except Exception as e: # Other potential errors
    print(f"\n⚠️ 無法自動開啟資料夾：{e}")
    print(f"📁 請手動開啟：{os.path.abspath(output_folder)}")
