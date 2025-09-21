import sqlite3
import os

def add_cod_to_database():
    """Add Call of Duty to the database with the correct path"""
    
    # The correct path we found
    cod_path = r"C:\XboxGames\Call of Duty\Content\cod.exe"
    
    # Verify the file exists
    if not os.path.exists(cod_path):
        print(f"❌ File not found: {cod_path}")
        return False
    
    print(f"✅ Found Call of Duty at: {cod_path}")
    
    # Connect to the database
    db_path = "iris_voice_learning.db"
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    try:
        # Add Call of Duty to detected_programs table
        print("📝 Adding Call of Duty to detected_programs table...")
        cursor.execute('''
            INSERT OR REPLACE INTO detected_programs (program_name, program_path, category, platform)
            VALUES (?, ?, ?, ?)
        ''', ("Call of Duty Black Ops 6", cod_path, "Game", "Xbox Game Pass"))
        
        # Add all the voice aliases
        print("🗣️ Adding voice aliases...")
        aliases = [
            "call of duty",
            "cod", 
            "black ops",
            "bo6",
            "cod black ops",
            "cod6",
            "call of duty black ops 6",
            "black ops 6"
        ]
        
        for alias in aliases:
            cursor.execute('''
                INSERT OR REPLACE INTO program_aliases (program_name, alias, weight)
                VALUES (?, ?, ?)
            ''', ("Call of Duty Black Ops 6", alias, 1.5))
            print(f"   • Added alias: '{alias}'")
        
        # Commit changes
        conn.commit()
        print("✅ Successfully added Call of Duty to database!")
        
        # Verify it was added
        cursor.execute("SELECT program_name, program_path FROM detected_programs WHERE program_name LIKE '%call of duty%'")
        result = cursor.fetchone()
        if result:
            print(f"✅ Verification: {result[0]} → {result[1]}")
            return True
        else:
            print("❌ Verification failed - not found in database")
            return False
            
    except Exception as e:
        print(f"❌ Error adding to database: {e}")
        return False
    
    finally:
        conn.close()

def update_training_script():
    """Update the training script to include C:\XboxGames path"""
    print("\n" + "=" * 60)
    print("🔧 Updating Training Script")
    print("=" * 60)
    
    # Read the current training.py file
    try:
        with open("training.py", "r", encoding="utf-8") as f:
            content = f.read()
        
        # Check if C:\XboxGames is already included
        if "C:\\\\XboxGames" in content or "C:/XboxGames" in content:
            print("✅ Training script already includes C:\\XboxGames")
            return True
        
        # Find the xbox_path line and add our path
        if 'xbox_path = "C:\\\\Program Files\\\\WindowsApps"' in content:
            # Replace the single path with multiple paths
            old_section = '''xbox_path = "C:\\\\Program Files\\\\WindowsApps"
            if os.path.exists(xbox_path):'''
            
            new_section = '''xbox_paths = [
                "C:\\\\Program Files\\\\WindowsApps",
                "C:\\\\XboxGames"
            ]
            for xbox_path in xbox_paths:
                if not os.path.exists(xbox_path):
                    continue'''
            
            content = content.replace(old_section, new_section)
            
            # Also need to fix the indentation for the rest of the Xbox detection
            content = content.replace('''                # Get accessible subdirectories
                try:
                    for item in os.listdir(xbox_path):''', '''                    # Get accessible subdirectories
                    try:
                        for item in os.listdir(xbox_path):''')
            
            # Write the updated content
            with open("training.py", "w", encoding="utf-8") as f:
                f.write(content)
            
            print("✅ Updated training.py to include C:\\XboxGames")
            return True
        else:
            print("⚠️ Couldn't find the right section to update in training.py")
            return False
            
    except FileNotFoundError:
        print("❌ training.py not found")
        return False
    except Exception as e:
        print(f"❌ Error updating training.py: {e}")
        return False

def test_database_fix():
    """Test if the database fix worked"""
    print("\n" + "=" * 60)
    print("🧪 Testing Database Fix")
    print("=" * 60)
    
    db_path = "iris_voice_learning.db"
    if not os.path.exists(db_path):
        print("❌ Database not found!")
        return False
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Check if COD is now in detected_programs
    cursor.execute("SELECT program_name, program_path, category, platform FROM detected_programs WHERE program_name LIKE '%call of duty%'")
    cod_programs = cursor.fetchall()
    
    if cod_programs:
        print("✅ Found Call of Duty in detected_programs:")
        for name, path, category, platform in cod_programs:
            print(f"   • {name}")
            print(f"     Path: {path}")
            print(f"     Category: {category}")
            print(f"     Platform: {platform}")
    else:
        print("❌ Call of Duty still not found in detected_programs")
        return False
    
    # Check aliases
    cursor.execute("SELECT alias FROM program_aliases WHERE program_name LIKE '%call of duty%'")
    aliases = cursor.fetchall()
    
    if aliases:
        print(f"\n✅ Found {len(aliases)} voice aliases:")
        for (alias,) in aliases:
            print(f"   • '{alias}'")
    else:
        print("❌ No voice aliases found")
    
    conn.close()
    return True

if __name__ == "__main__":
    print("🔧 Fixing Call of Duty Database Issue")
    print("=" * 60)
    
    # Step 1: Add COD to the database manually
    if add_cod_to_database():
        print("\n✅ Step 1 Complete: Added Call of Duty to database")
        
        # Step 2: Test the fix
        if test_database_fix():
            print("\n✅ Step 2 Complete: Database fix verified")
            
            # Step 3: Update training script for future runs
            print("\n🔄 Step 3: Updating training script...")
            if update_training_script():
                print("✅ Step 3 Complete: Training script updated")
            
            print("\n" + "=" * 60)
            print("🎉 SUCCESS! Call of Duty should now work!")
            print("🎤 Try saying: 'Hey Iris' → 'Open Call of Duty'")
            print("🎮 Or just: 'Hey Iris' → 'Call of Duty'")
            print("=" * 60)
        else:
            print("❌ Database fix verification failed")
    else:
        print("❌ Failed to add Call of Duty to database")