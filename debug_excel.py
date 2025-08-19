#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script debug để kiểm tra file Excel và tìm nguyên nhân tại sao chỉ nhận 7/10 link
"""

import pandas as pd
import os

def debug_excel_file(file_path):
    """Debug file Excel để tìm nguyên nhân mất link"""
    
    print(f"🔍 DEBUG FILE EXCEL: {file_path}")
    print("=" * 60)
    
    try:
        # Đọc file Excel
        df = pd.read_excel(file_path)
        
        print(f"📊 Thông tin file:")
        print(f"   - Tổng số cột: {len(df.columns)}")
        print(f"   - Tổng số dòng: {len(df)}")
        print(f"   - Tên cột: {list(df.columns)}")
        print()
        
        # Kiểm tra từng dòng
        print("📋 KIỂM TRA TỪNG DÒNG:")
        print("-" * 60)
        
        valid_count = 0
        skipped_count = 0
        
        for i in range(len(df)):
            # Lấy dữ liệu từ cột A và B
            col_a = df.iloc[i, 0] if len(df.columns) > 0 else None
            col_b = df.iloc[i, 1] if len(df.columns) > 1 else None
            
            # Kiểm tra dữ liệu
            code_str = str(col_a).strip() if pd.notna(col_a) else ""
            link_str = str(col_b).strip() if pd.notna(col_b) else ""
            
            print(f"Dòng {i+1:2d}:")
            print(f"   Cột A (Mã): '{col_a}' -> '{code_str}'")
            print(f"   Cột B (Link): '{col_b}' -> '{link_str}'")
            
            # Phân tích lý do bỏ qua
            if not code_str and not link_str:
                print(f"   ❌ BỎ QUA: Dòng trống hoàn toàn")
                skipped_count += 1
            elif code_str and not link_str:
                print(f"   ❌ BỎ QUA: Có mã nhưng không có link")
                skipped_count += 1
            elif not code_str and link_str:
                print(f"   ⚠️  TỰ TẠO MÃ: Có link nhưng không có mã")
                valid_count += 1
            else:
                print(f"   ✅ HỢP LỆ: Có cả mã và link")
                valid_count += 1
            
            print()
        
        print("📊 KẾT QUẢ PHÂN TÍCH:")
        print(f"   ✅ Hợp lệ: {valid_count} dòng")
        print(f"   ❌ Bỏ qua: {skipped_count} dòng")
        print(f"   📋 Tổng cộng: {len(df)} dòng")
        
        # Kiểm tra các vấn đề thường gặp
        print("\n🔍 KIỂM TRA VẤN ĐỀ THƯỜNG GẶP:")
        print("-" * 40)
        
        # 1. Kiểm tra dòng trống ở cuối
        if len(df) > 0:
            last_row_a = df.iloc[-1, 0] if len(df.columns) > 0 else None
            last_row_b = df.iloc[-1, 1] if len(df.columns) > 1 else None
            if pd.isna(last_row_a) and pd.isna(last_row_b):
                print("⚠️  Phát hiện: Dòng cuối trống (có thể do Excel tự thêm)")
        
        # 2. Kiểm tra dòng header
        first_row_a = df.iloc[0, 0] if len(df.columns) > 0 else None
        first_row_b = df.iloc[0, 1] if len(df.columns) > 1 else None
        if isinstance(first_row_a, str) and "mã" in first_row_a.lower():
            print("⚠️  Phát hiện: Dòng đầu có thể là header")
        
        # 3. Kiểm tra dữ liệu ẩn
        for i in range(len(df)):
            for j in range(min(2, len(df.columns))):
                cell_value = df.iloc[i, j]
                if isinstance(cell_value, str) and cell_value.strip() == "":
                    print(f"⚠️  Phát hiện: Ô trống ở dòng {i+1}, cột {j+1}")
        
        print("\n💡 GỢI Ý KHẮC PHỤC:")
        print("1. Xóa các dòng trống hoàn toàn")
        print("2. Đảm bảo mỗi dòng có ít nhất 1 link")
        print("3. Kiểm tra không có dòng header")
        print("4. Xóa các ký tự ẩn (space, tab) thừa")
        
    except Exception as e:
        print(f"❌ Lỗi khi đọc file Excel: {str(e)}")

def main():
    """Hàm chính"""
    print("🔍 DEBUG EXCEL - KIỂM TRA TẠI SAO CHỈ NHẬN 7/10 LINK")
    print("=" * 70)
    
    # Tìm file Excel trong thư mục hiện tại
    excel_files = []
    for file in os.listdir('.'):
        if file.endswith(('.xlsx', '.xls')):
            excel_files.append(file)
    
    if not excel_files:
        print("❌ Không tìm thấy file Excel nào trong thư mục hiện tại!")
        print("📁 Các file hiện có:")
        for file in os.listdir('.'):
            print(f"   - {file}")
        return
    
    print(f"📁 Tìm thấy {len(excel_files)} file Excel:")
    for i, file in enumerate(excel_files, 1):
        print(f"   {i}. {file}")
    
    # Debug file đầu tiên (hoặc yêu cầu người dùng chọn)
    if len(excel_files) == 1:
        debug_excel_file(excel_files[0])
    else:
        print(f"\n🔍 Tự động debug file đầu tiên: {excel_files[0]}")
        debug_excel_file(excel_files[0])

if __name__ == "__main__":
    main()
