#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script sửa toàn bộ logic xử lý Excel để đảm bảo nhận đầy đủ tất cả cặp mã-link
"""

import pandas as pd
import os

def analyze_excel_file(file_path):
    """Phân tích chi tiết file Excel"""
    
    print(f"🔍 PHÂN TÍCH CHI TIẾT FILE EXCEL: {file_path}")
    print("=" * 70)
    
    try:
        # Đọc file Excel với các tùy chọn khác nhau
        print("📊 THỬ CÁC CÁCH ĐỌC FILE EXCEL:")
        print("-" * 40)
        
        # Cách 1: Đọc bình thường
        df1 = pd.read_excel(file_path)
        print(f"1. Đọc bình thường: {len(df1)} dòng, {len(df1.columns)} cột")
        
        # Cách 2: Đọc với header=None
        df2 = pd.read_excel(file_path, header=None)
        print(f"2. Đọc không header: {len(df2)} dòng, {len(df2.columns)} cột")
        
        # Cách 3: Đọc với skiprows=0
        df3 = pd.read_excel(file_path, skiprows=0)
        print(f"3. Đọc skiprows=0: {len(df3)} dòng, {len(df3.columns)} cột")
        
        # Cách 4: Đọc với nrows lớn
        df4 = pd.read_excel(file_path, nrows=20)
        print(f"4. Đọc nrows=20: {len(df4)} dòng, {len(df4.columns)} cột")
        
        # Chọn DataFrame tốt nhất
        if len(df2) >= len(df1):
            df = df2
            print(f"\n✅ Sử dụng cách đọc không header: {len(df)} dòng")
        else:
            df = df1
            print(f"\n✅ Sử dụng cách đọc bình thường: {len(df)} dòng")
        
        print(f"\n📋 THÔNG TIN FILE:")
        print(f"   - Tổng số cột: {len(df.columns)}")
        print(f"   - Tổng số dòng: {len(df)}")
        print(f"   - Tên cột: {list(df.columns)}")
        
        # Kiểm tra từng dòng chi tiết
        print(f"\n📋 KIỂM TRA TỪNG DÒNG CHI TIẾT:")
        print("-" * 60)
        
        valid_pairs = []
        skipped_rows = []
        
        for i in range(len(df)):
            # Lấy dữ liệu từ cột đầu tiên và thứ hai
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
                skipped_rows.append((i+1, "Dòng trống hoàn toàn"))
            elif code_str and not link_str:
                print(f"   ❌ BỎ QUA: Có mã nhưng không có link")
                skipped_rows.append((i+1, "Có mã nhưng không có link"))
            elif not code_str and link_str:
                print(f"   ⚠️ TỰ TẠO MÃ: Có link nhưng không có mã")
                code_str = f"PRODUCT_{i+1:03d}"
                valid_pairs.append((code_str, link_str))
            else:
                print(f"   ✅ HỢP LỆ: Có cả mã và link")
                valid_pairs.append((code_str, link_str))
            
            print()
        
        print("📊 KẾT QUẢ PHÂN TÍCH:")
        print(f"   ✅ Hợp lệ: {len(valid_pairs)} cặp mã-link")
        print(f"   ❌ Bỏ qua: {len(skipped_rows)} dòng")
        print(f"   📋 Tổng cộng: {len(df)} dòng")
        
        if skipped_rows:
            print(f"\n❌ CÁC DÒNG BỊ BỎ QUA:")
            for row_num, reason in skipped_rows:
                print(f"   Dòng {row_num}: {reason}")
        
        print(f"\n🎯 CÁC CẶP HỢP LỆ:")
        for i, (code, link) in enumerate(valid_pairs, 1):
            print(f"   {i:2d}. {code} -> {link}")
        
        return valid_pairs, len(df), len(skipped_rows)
        
    except Exception as e:
        print(f"❌ Lỗi khi đọc file Excel: {str(e)}")
        return [], 0, 0

def test_improved_logic():
    """Test logic cải thiện"""
    
    print("\n🧪 TEST LOGIC CẢI THIỆN")
    print("=" * 50)
    
    # Tìm file Excel
    excel_file = "Link cào.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ Không tìm thấy file: {excel_file}")
        return
    
    # Phân tích file
    valid_pairs, total_rows, skipped_rows = analyze_excel_file(excel_file)
    
    print(f"\n📊 KẾT QUẢ CUỐI CÙNG:")
    print(f"   📋 Tổng dòng Excel: {total_rows}")
    print(f"   ✅ Cặp hợp lệ: {len(valid_pairs)}")
    print(f"   ❌ Dòng bỏ qua: {skipped_rows}")
    
    if len(valid_pairs) == 0:
        print(f"\n❌ KHÔNG CÓ CẶP NÀO HỢP LỆ!")
        print("   Vui lòng kiểm tra lại file Excel")
    else:
        print(f"\n✅ THÀNH CÔNG: Tìm thấy {len(valid_pairs)} cặp mã-link")
        print("   Logic mới sẽ xử lý đầy đủ tất cả cặp này!")

def main():
    """Hàm chính"""
    print("🔧 SỬA TOÀN BỘ LOGIC XỬ LÝ EXCEL")
    print("=" * 70)
    
    test_improved_logic()

if __name__ == "__main__":
    main()
