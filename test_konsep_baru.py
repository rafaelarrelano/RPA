#!/usr/bin/env python3
"""
Test konsep baru: email & u2c terpisah untuk items lewat limit
"""

from main import StockDiff

def test_stockdiff_flags():
    """Test bahwa StockDiff dapat membuat dengan flag skip"""
    
    # Test 1: Create normal item
    item1 = StockDiff(
        param="FSTKGD",
        plant="4503",
        sloc="WH01",
        posting_date="01.04.2026",
        material="377001",
        qty_matrix=100,
        qty_sap=101,
        diff=1,
        status=0,
        mvt_type="918",
        qty_adjust=1.0
    )
    assert item1.should_skip_email == False
    assert item1.should_skip_u2c == False
    print("✓ Test 1: Item normal OK")
    
    # Test 2: Create dengan flag skip
    item2 = StockDiff(
        param="FSTKGD",
        plant="4503",
        sloc="WH01",
        posting_date="01.04.2026",
        material="377002",
        qty_matrix=100,
        qty_sap=101,
        diff=1,
        status=0,
        should_skip_email=True,
        should_skip_u2c=True
    )
    assert item2.should_skip_email == True
    assert item2.should_skip_u2c == True
    print("✓ Test 2: Item dengan skip flag OK")
    
    # Test 3: Modify flag
    item1.should_skip_email = True
    item1.should_skip_u2c = True
    assert item1.should_skip_email == True
    assert item1.should_skip_u2c == True
    print("✓ Test 3: Modify flag OK")


def test_filter_skip_email():
    """Test filter untuk email"""
    items = [
        StockDiff("FSTKGD","4503","WH01","01.04.2026","377001",100,101,1,0,"918",1.0),
        StockDiff("FSTKGD","4503","WH01","01.04.2026","377002",100,101,1,0,"918",1.0,
                 should_skip_email=True),
        StockDiff("FSTKGD","4503","WH01","01.04.2026","377003",100,101,1,0,"918",1.0),
    ]
    
    # Simulasi: tandai item ke-2 (index 1) sebagai skip
    items[1].should_skip_email = True
    
    # Filter untuk email
    items_to_send = [i for i in items if not getattr(i, 'should_skip_email', False)]
    items_to_skip = [i for i in items if getattr(i, 'should_skip_email', False)]
    
    assert len(items_to_send) == 2
    assert len(items_to_skip) == 1
    assert items_to_skip[0].material == "377002"
    print("✓ Test 4: Filter email OK")


def test_filter_skip_u2c():
    """Test filter untuk U2C upload"""
    items = [
        StockDiff("FSTKGD","4503","WH01","01.04.2026","377001",100,101,1,0,"918",1.0),
        StockDiff("FSTKGD","4503","WH01","01.04.2026","377002",100,101,1,0,"918",1.0,
                 should_skip_u2c=True),
        StockDiff("FSTKGD","4503","WH01","01.04.2026","377003",100,101,1,0,"918",1.0),
    ]
    
    # Filter untuk U2C
    items_to_upload = [i for i in items if not getattr(i, 'should_skip_u2c', False)]
    items_to_skip = [i for i in items if getattr(i, 'should_skip_u2c', False)]
    
    assert len(items_to_upload) == 2
    assert len(items_to_skip) == 1
    assert items_to_skip[0].material == "377002"
    print("✓ Test 5: Filter U2C OK")


def test_combined_scenario():
    """Test skenario lengkap dengan multiple plants"""
    
    # Plant 4503 - ada item skip
    plant_4503_items = [
        StockDiff("FSTKGD","4503","WH01","01.04.2026","377001",100,101,1,0,"918",1.0),
        StockDiff("FSTKGD","4503","WH01","01.04.2026","377002",100,101,1,0,"918",1.0),
    ]
    plant_4503_items[1].should_skip_email = True
    plant_4503_items[1].should_skip_u2c = True
    
    # Plant 4507 - semua OK
    plant_4507_items = [
        StockDiff("FSTKGD","4507","WH01","01.04.2026","378001",100,101,1,0,"918",1.0),
    ]
    
    items_per_plant = {
        "4503": plant_4503_items,
        "4507": plant_4507_items,
    }
    
    # Filter email
    items_email = {}
    for plant, items in items_per_plant.items():
        items_to_send = [i for i in items if not getattr(i, 'should_skip_email', False)]
        if items_to_send:
            items_email[plant] = items_to_send
    
    assert len(items_email["4503"]) == 1
    assert len(items_email["4507"]) == 1
    print("✓ Test 6: Combined scenario OK")


if __name__ == "__main__":
    print("Testing konsep baru: email & u2c terpisah...\n")
    test_stockdiff_flags()
    test_filter_skip_email()
    test_filter_skip_u2c()
    test_combined_scenario()
    print("\n✓ Semua test berhasil!")
