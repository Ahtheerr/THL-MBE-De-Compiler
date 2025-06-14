# mbe_decompiler.py
import struct
import openpyxl
import os
import argparse
import re
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# --- MUDANÇA AQUI: Adicionado novo tipo de coluna ---
COL_TYPE_INT, COL_TYPE_STRING, COL_TYPE_STRINGID, COL_TYPE_INTID = 0x2, 0x7, 0x8, 0x9
COL_TYPE_NAMES = {
    COL_TYPE_INT: "Int",
    COL_TYPE_STRING: "String",
    COL_TYPE_STRINGID: "StringID",
    COL_TYPE_INTID: "IntID", # Mapeamento para o nome
}

COLOR_COMMAND_PATTERN = re.compile(r'\{fc\(([a-fA-F0-9]{6})\)(.*?)\}', re.DOTALL)

def create_visual_rich_text(raw_text):
    if not isinstance(raw_text, str) or not COLOR_COMMAND_PATTERN.search(raw_text): return raw_text
    default_font = InlineFont()
    text_blocks = []
    last_index = 0
    for match in COLOR_COMMAND_PATTERN.finditer(raw_text):
        uncolored_prefix_end = match.start(2)
        if uncolored_prefix_end > last_index:
            text_blocks.append(TextBlock(default_font, raw_text[last_index:uncolored_prefix_end]))
        hex_color = match.group(1)
        inner_text = match.group(2)
        if inner_text:
            colored_font = InlineFont(color=f"00{hex_color}")
            text_blocks.append(TextBlock(colored_font, inner_text))
        last_index = match.end(2)
    if last_index < len(raw_text):
        text_blocks.append(TextBlock(default_font, raw_text[last_index:]))
    return CellRichText(text_blocks)

def decompile_mbe(mbe_path, xlsx_path):
    print(f"Iniciando decompilação de '{mbe_path}'...")
    try:
        with open(mbe_path, 'rb') as f:
            f.seek(4); f.read(4)
            tab_name_size = struct.unpack('<I', f.read(4))[0]
            tab_name = f.read(tab_name_size).rstrip(b'\x00').decode('utf-8', 'ignore')
            num_columns = struct.unpack('<I', f.read(4))[0]
            column_types = [struct.unpack('<I', f.read(4))[0] for _ in range(num_columns)]
            row_size_from_file, num_rows = struct.unpack('<II', f.read(8))
            
            expa_data_start_offset = f.tell()
            if num_rows > 0 and column_types and column_types[0] == COL_TYPE_INT:
                current_pos = f.tell()
                first_val_bytes = f.read(4)
                if first_val_bytes and struct.unpack('<i', first_val_bytes)[0] == 0:
                    expa_data_start_offset += 4
                f.seek(current_pos)
            
            f.seek(expa_data_start_offset)
            expa_data_blob = f.read(num_rows * row_size_from_file)
            
            magic_chnk = f.read(4)
            if magic_chnk != b'CHNK': raise ValueError("Seção CHNK não encontrada.")
            num_strings = struct.unpack('<I', f.read(4))[0]
            string_map = {}
            for _ in range(num_strings):
                target_offset, string_size = struct.unpack('<II', f.read(8))
                string_map[target_offset] = f.read(string_size).rstrip(b'\x00').decode('utf-8', 'ignore')

            all_rows_data = []
            for i in range(num_rows):
                row_data, offset_in_row = [], 0
                for col_type in column_types:
                    alignment = 8 if col_type in [COL_TYPE_STRING, COL_TYPE_STRINGID] else 4
                    padding = (alignment - (offset_in_row % alignment)) % alignment
                    offset_in_row += padding
                    data_start_in_blob = (i * row_size_from_file) + offset_in_row
                    
                    # --- MUDANÇA AQUI: Trata Int e IntID da mesma forma ---
                    if col_type in [COL_TYPE_INT, COL_TYPE_INTID]:
                        row_data.append(struct.unpack('<i', expa_data_blob[data_start_in_blob:data_start_in_blob+4])[0])
                        offset_in_row += 4
                    else: # String ou StringID
                        row_data.append(string_map.get(expa_data_start_offset + data_start_in_blob, ""))
                        offset_in_row += 8
                all_rows_data.append(row_data)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = tab_name
            ws.append([f"{COL_TYPE_NAMES.get(ct, 'Inválido')} ({i+1})" for i, ct in enumerate(column_types)])
            for row_data_list in all_rows_data:
                ws.append([create_visual_rich_text(val) for val in row_data_list])

            wb.save(xlsx_path)
            print("\nDecompilação concluída com sucesso!")
    except Exception as e:
        print(f"Ocorreu um erro durante a decompilação: {e}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Descompila .MBE para .XLSX com Rich Text visual.")
    parser.add_argument("input_file"); parser.add_argument("-o", "--output")
    args = parser.parse_args()
    decompile_mbe(args.input_file, args.output or f"{os.path.splitext(os.path.basename(args.input_file))[0]}.xlsx")