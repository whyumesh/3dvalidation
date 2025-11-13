# Clear existing data rows (between header and total)
        for r in range(data_start_row, total_row):
            for c in range(1, ws.max_column + 1):
                try:
                    cell = ws.cell(row=r, column=c)
                    cell.value = None
                except:
                    pass

        def copy_row_style(src_row_idx, dst_row_idx):
            """Copy formatting from source row to destination row"""
            for c in range(1, ws.max_column + 1):
                try:
                    src = ws.cell(row=src_row_idx, column=c)
                    dst = ws.cell(row=dst_row_idx, column=c)
                    
                    if src.font:
                        dst.font = copy_style(src.font)
                    if src.alignment:
                        dst.alignment = copy_style(src.alignment)
                    if src.border:
                        dst.border = copy_style(src.border)
                    if src.fill:
                        dst.fill = copy_style(src.fill)
                    dst.number_format = src.number_format
                except:
                    pass

        # Write actual Division data row (data_start_row is the first row after header)
        copy_row_style(total_row, data_start_row)
        
        for col_name, col_idx in column_mapping.items():
            if col_name in summary_df.columns:
                value = summary_df.iloc[0][col_name]
                
                try:
                    cell = ws.cell(row=data_start_row, column=col_idx)
                    cell.value = value
                    
                    if isinstance(value, (int, float)) and not pd.isna(value):
                        cell.number_format = '0'
                        cell.font = Font(name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        cell.font = Font(name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                except Exception as e:
                    print(f"   Warning: Could not set value for column {col_name}: {e}")
        
        # Write data to Total row with same values (since there's only one Division per file)
        copy_row_style(total_row, total_row)
        
        # Set "Total" text in first column
        ws.cell(row=total_row, column=1, value="Total")
        
        for col_name, col_idx in column_mapping.items():
            if col_name in summary_df.columns:
                value = summary_df.iloc[0][col_name]
                
                try:
                    cell = ws.cell(row=total_row, column=col_idx)
                    cell.value = value
                    
                    if isinstance(value, (int, float)) and not pd.isna(value):
                        cell.number_format = '0'
                        cell.font = Font(bold=True, name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        cell.font = Font(bold=True, name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                except Exception as e:
                    print(f"   Warning: Could not set value for column {col_name}: {e}")
