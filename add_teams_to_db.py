import sqlite3
import os

def add_teams_to_database():
    """Add the found Microsoft Teams to the database"""
    print("📝 Adding Microsoft Teams to Database")
    print("=" * 60)
    
    # The main Teams executable (the most important one)
    teams_exe = r"C:\Users\Ahmed\AppData\Local\Microsoft\WindowsApps\ms-teams.exe"
    
    # Verify it exists
    if not os.path.exists(teams_exe):
        print(f"❌ Teams executable not found: {teams_exe}")
        return False
    
    print(f"✅ Found Microsoft Teams at: {teams_exe}")
    
    # Connect to database
    db_path = "iris_voice_learning.db"
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    try:
        # Add Microsoft Teams to detected_programs table
        print("📝 Adding to detected_programs table...")
        cursor.execute('''
            INSERT OR REPLACE INTO detected_programs (program_name, program_path, category, platform)
            VALUES (?, ?, ?, ?)
        ''', ("Microsoft Teams", teams_exe, "Application", "Microsoft Store"))
        
        # Add comprehensive voice aliases
        print("🗣️ Adding voice aliases...")
        teams_aliases = [
            "microsoft teams",
            "teams",
            "ms teams",
            "team", 
            "teams app",
            "meeting",
            "video call",
            "conference",
            "teams meeting",
            "join meeting",
            "video chat"
        ]
        
        for alias in teams_aliases:
            cursor.execute('''
                INSERT OR REPLACE INTO program_aliases (program_name, alias, weight)
                VALUES (?, ?, ?)
            ''', ("Microsoft Teams", alias, 1.5))
            print(f"   • Added: '{alias}'")
        
        # Commit changes
        conn.commit()
        print("\n✅ Microsoft Teams successfully added to database!")
        
        # Verify the addition
        cursor.execute("SELECT program_name, program_path FROM detected_programs WHERE program_name = 'Microsoft Teams'")
        result = cursor.fetchone()
        
        if result:
            print(f"✅ Verification successful: {result[0]} → {result[1]}")
            return True
        else:
            print("❌ Verification failed")
            return False
            
    except Exception as e:
        print(f"❌ Error adding to database: {e}")
        return False
    
    finally:
        conn.close()

def test_teams_aliases():
    """Test that the aliases work correctly"""
    print("\n🧪 Testing Teams Aliases")
    print("=" * 30)
    
    db_path = "iris_voice_learning.db"
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Get all Teams aliases
    cursor.execute("SELECT alias FROM program_aliases WHERE program_name = 'Microsoft Teams'")
    aliases = cursor.fetchall()
    
    print(f"Found {len(aliases)} aliases for Microsoft Teams:")
    for (alias,) in aliases:
        print(f"   🗣️ '{alias}'")
    
    conn.close()
    
    print("\n✅ You can now use any of these voice commands:")
    print("   • 'Hey Iris' → 'Open Microsoft Teams'")
    print("   • 'Hey Iris' → 'Teams'")
    print("   • 'Hey Iris' → 'Open Teams'") 
    print("   • 'Hey Iris' → 'Join meeting'")
    print("   • 'Hey Iris' → 'Video call'")

def fix_speech_recognition_issues():
    """Add common misheard versions of 'teams' commands"""
    print("\n🔧 Adding Speech Recognition Error Fixes")
    print("=" * 45)
    
    db_path = "iris_voice_learning.db"
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Common misheard versions
    misheard_aliases = [
        # Common speech recognition errors for "teams"
        "tims",           # teams → tims
        "teens",          # teams → teens  
        "dreams",         # teams → dreams
        "beams",          # teams → beams
        "seams",          # teams → seams
        "steams",         # teams → steams
        
        # Misheard "microsoft teams"
        "microsoft tims",
        "microsoft teens", 
        "micro soft teams",
        "microsof teams",
        
        # Common context words that might be misheard
        "open tims",
        "start teams",
        "launch teams"
    ]
    
    print("Adding misheard aliases to improve recognition:")
    for alias in misheard_aliases:
        try:
            cursor.execute('''
                INSERT OR REPLACE INTO program_aliases (program_name, alias, weight)
                VALUES (?, ?, ?)
            ''', ("Microsoft Teams", alias, 1.2))  # Slightly lower weight for misheard versions
            print(f"   • Added: '{alias}'")
        except Exception as e:
            print(f"   ⚠️ Failed to add '{alias}': {e}")
    
    conn.commit()
    conn.close()
    
    print("✅ Added speech recognition error fixes!")

def show_final_summary():
    """Show final summary and test suggestions"""
    print("\n" + "=" * 60)
    print("🎉 MICROSOFT TEAMS SETUP COMPLETE!")
    print("=" * 60)
    
    print("✅ Microsoft Teams has been added to your database")
    print("✅ Voice aliases have been configured") 
    print("✅ Speech recognition error fixes added")
    
    print("\n🎤 Test Commands:")
    print("   1. Say: 'Hey Iris' → wait for response → 'Open Teams'")
    print("   2. Say: 'Hey Iris' → wait for response → 'Microsoft Teams'")
    print("   3. Say: 'Hey Iris' → wait for response → 'Join meeting'")
    print("   4. Say: 'Hey Iris' → wait for response → 'Teams'")
    
    print("\n🔧 Path Details:")
    print(f"   • Teams Executable: C:\\Users\\Ahmed\\AppData\\Local\\Microsoft\\WindowsApps\\ms-teams.exe")
    print(f"   • Database: iris_voice_learning.db")
    print(f"   • Category: Application")
    print(f"   • Platform: Microsoft Store")
    
    print("\n💡 Troubleshooting:")
    print("   • If Teams doesn't open, make sure it's properly installed")
    print("   • Try running your main.py and test the voice commands")
    print("   • Check the debug output to see if Teams is being found")
    
    print("=" * 60)

if __name__ == "__main__":
    # Step 1: Add Teams to database
    if add_teams_to_database():
        # Step 2: Test the aliases
        test_teams_aliases()
        
        # Step 3: Add speech recognition fixes
        fix_speech_recognition_issues()
        
        # Step 4: Show final summary
        show_final_summary()
        
        print("\n🚀 Ready to test! Run your main.py and try the voice commands!")
    else:
        print("❌ Failed to add Microsoft Teams to database")