#!/usr/bin/env python3
"""
Test script to verify the specialty-based configuration system
"""

from src.word_generator import WordGenerator

def test_config_system():
    """Test the configuration system with different levels and specialties"""
    wg = WordGenerator()
    
    print("Testing configuration system:")
    print("=" * 50)
    
    # Test different level and specialty combinations
    test_cases = [
        ("100", "comm"),
        ("100", "comp"), 
        ("100", "pow"),
        ("200", "comm"),
        ("200", "comp"),
        ("200", "pow"),
        ("300", "comm"),
        ("300", "comp"),
        ("300", "pow"),
        ("400", "comm"),
        ("400", "comp"),
        ("400", "pow"),
    ]
    
    for level, specialty in test_cases:
        config = wg._get_level_config(level, specialty)
        header_config = config["header"]
        
        print(f"\nLevel {level}, Specialty {specialty}:")
        print(f"  Department: {header_config['department_prefix']}")
        print(f"  Level Prefix: {header_config['level_prefix']}")
        print(f"  Schedule Template: {header_config['schedule_template']}")
        
        # Check if program manager is present (should be for levels 100, 200)
        footer_config = config["footer"]
        has_program_manager = "program_manager_title" in footer_config
        print(f"  Has Program Manager: {has_program_manager}")

if __name__ == "__main__":
    test_config_system()
