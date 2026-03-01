    print(f"フルネーム: {fn}")
    print(f"別名マップ: {am}")
    
    rows, chars = parse_lines_to_kouban(test_lines)
    print(f"\n抽出キャラクター: {chars}")
    print(f"\nシーン数: {len(rows)}")
    
    for row in rows:
        print(f"\nシーン{row['scene_no']}: {row['location']} ({row['D/N']})")
        print(f"  要約: {row['summary'][:40]}...")
        present = [c for c in chars if row.get(c) == "◯"]
        print(f"  登場: {present}")
    
    print("\n=== CSV出力 ===")
    csv_str = convert_to_csv_string(rows, chars)
    print(csv_str[:500])