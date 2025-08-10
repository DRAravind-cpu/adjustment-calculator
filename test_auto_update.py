#!/usr/bin/env python3
"""
Test script for the auto-update system
Verifies all components work correctly
"""

import os
import sys
import json
import tempfile
import shutil
from pathlib import Path
import unittest
from unittest.mock import Mock, patch, MagicMock

# Add current directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

try:
    from auto_updater import AutoUpdater, initialize_updater
    UPDATER_AVAILABLE = True
except ImportError as e:
    print(f"❌ Auto-updater import failed: {e}")
    UPDATER_AVAILABLE = False

class TestAutoUpdater(unittest.TestCase):
    """Test cases for the auto-updater system"""
    
    def setUp(self):
        """Set up test environment"""
        if not UPDATER_AVAILABLE:
            self.skipTest("Auto-updater not available")
        
        self.test_dir = Path(tempfile.mkdtemp())
        self.original_dir = os.getcwd()
        os.chdir(self.test_dir)
        
        # Create test version file
        self.version_data = {
            "version": "1.0.0",
            "release_date": "2024-08-10",
            "build_number": "001"
        }
        
        with open("version.json", "w") as f:
            json.dump(self.version_data, f)
        
        self.updater = AutoUpdater("1.0.0", "Test App")
    
    def tearDown(self):
        """Clean up test environment"""
        os.chdir(self.original_dir)
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def test_updater_initialization(self):
        """Test updater initialization"""
        self.assertEqual(self.updater.current_version, "1.0.0")
        self.assertEqual(self.updater.app_name, "Test App")
        self.assertTrue(self.updater.config_file.name == "update_config.json")
    
    def test_config_loading(self):
        """Test configuration loading and saving"""
        # Test default config
        config = self.updater.load_config()
        self.assertTrue(config["auto_check"])
        self.assertEqual(config["update_channel"], "stable")
        
        # Test config saving
        config["auto_check"] = False
        self.updater.config = config
        self.updater.save_config()
        
        # Reload and verify
        new_updater = AutoUpdater("1.0.0", "Test App")
        self.assertFalse(new_updater.config["auto_check"])
    
    def test_version_comparison(self):
        """Test version comparison logic"""
        # Test normal versions
        self.assertEqual(self.updater.compare_versions("1.0.0", "1.0.1"), -1)
        self.assertEqual(self.updater.compare_versions("1.1.0", "1.0.0"), 1)
        self.assertEqual(self.updater.compare_versions("1.0.0", "1.0.0"), 0)
        
        # Test dev versions
        self.assertEqual(self.updater.compare_versions("1.0.0", "dev-abc123"), -1)
    
    @patch('requests.get')
    def test_internet_connection_check(self, mock_get):
        """Test internet connection checking"""
        # Test successful connection
        mock_response = Mock()
        mock_response.status_code = 200
        mock_get.return_value = mock_response
        
        self.assertTrue(self.updater.check_internet_connection())
        
        # Test failed connection
        mock_get.side_effect = Exception("Connection failed")
        self.assertFalse(self.updater.check_internet_connection())
    
    @patch('requests.get')
    def test_update_check(self, mock_get):
        """Test update checking"""
        # Mock successful API response
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "tag_name": "v1.1.0",
            "name": "Version 1.1.0",
            "body": "New features and bug fixes",
            "zipball_url": "https://github.com/test/repo/archive/v1.1.0.zip",
            "published_at": "2024-08-10T12:00:00Z",
            "assets": []
        }
        mock_get.return_value = mock_response
        
        update_info = self.updater.get_latest_version_info()
        
        self.assertIsNotNone(update_info)
        self.assertEqual(update_info["version"], "1.1.0")
        self.assertEqual(update_info["name"], "Version 1.1.0")
    
    def test_should_check_updates(self):
        """Test update check timing logic"""
        # Should check on first run
        self.assertTrue(self.updater.should_check_for_updates())
        
        # Disable auto-check
        self.updater.config["auto_check"] = False
        self.assertFalse(self.updater.should_check_for_updates())
    
    def test_backup_creation(self):
        """Test backup functionality"""
        # Create some test files
        test_file = self.test_dir / "test_app.py"
        test_file.write_text("print('Hello World')")
        
        # Test backup creation
        result = self.updater.backup_current_version()
        self.assertTrue(result)
        
        # Check backup directory exists
        backup_dir = self.test_dir / "backups"
        self.assertTrue(backup_dir.exists())
        
        # Check backup file exists
        backup_files = list(backup_dir.glob("backup_*.zip"))
        self.assertTrue(len(backup_files) > 0)

class TestReleaseManager(unittest.TestCase):
    """Test cases for the release manager"""
    
    def setUp(self):
        """Set up test environment"""
        self.test_dir = Path(tempfile.mkdtemp())
        self.original_dir = os.getcwd()
        os.chdir(self.test_dir)
        
        # Import here to avoid issues if file doesn't exist
        try:
            from release_manager import ReleaseManager
            self.manager = ReleaseManager()
        except ImportError:
            self.skipTest("Release manager not available")
    
    def tearDown(self):
        """Clean up test environment"""
        os.chdir(self.original_dir)
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def test_version_increment(self):
        """Test version incrementing"""
        # Create initial version file
        version_data = {
            "version": "1.0.0",
            "release_date": "2024-08-10",
            "build_number": "001"
        }
        
        with open("version.json", "w") as f:
            json.dump(version_data, f)
        
        # Test patch increment
        new_version, new_data = self.manager.increment_version("patch")
        self.assertEqual(new_version, "1.0.1")
        
        # Test minor increment
        new_version, new_data = self.manager.increment_version("minor")
        self.assertEqual(new_version, "1.1.0")
        
        # Test major increment
        new_version, new_data = self.manager.increment_version("major")
        self.assertEqual(new_version, "2.0.0")

def run_integration_tests():
    """Run integration tests with real components"""
    print("🧪 Running Auto-Update Integration Tests")
    print("=" * 50)
    
    # Test 1: Updater initialization
    print("1. Testing updater initialization...")
    try:
        updater = initialize_updater("1.0.0")
        print("   ✅ Updater initialized successfully")
    except Exception as e:
        print(f"   ❌ Updater initialization failed: {e}")
        return False
    
    # Test 2: Configuration management
    print("2. Testing configuration management...")
    try:
        original_config = updater.config.copy()
        updater.config["auto_check"] = False
        updater.save_config()
        
        new_updater = initialize_updater("1.0.0")
        if not new_updater.config["auto_check"]:
            print("   ✅ Configuration persistence works")
        else:
            print("   ❌ Configuration not persisted")
        
        # Restore original config
        updater.config = original_config
        updater.save_config()
        
    except Exception as e:
        print(f"   ❌ Configuration test failed: {e}")
    
    # Test 3: Internet connectivity
    print("3. Testing internet connectivity...")
    try:
        has_internet = updater.check_internet_connection()
        if has_internet:
            print("   ✅ Internet connection available")
        else:
            print("   ⚠️  No internet connection (expected in offline environments)")
    except Exception as e:
        print(f"   ❌ Internet check failed: {e}")
    
    # Test 4: Version comparison
    print("4. Testing version comparison...")
    try:
        test_cases = [
            ("1.0.0", "1.0.1", -1),
            ("1.1.0", "1.0.0", 1),
            ("1.0.0", "1.0.0", 0),
            ("1.0.0", "dev-abc123", -1)
        ]
        
        all_passed = True
        for v1, v2, expected in test_cases:
            result = updater.compare_versions(v1, v2)
            if result != expected:
                print(f"   ❌ Version comparison failed: {v1} vs {v2}")
                all_passed = False
        
        if all_passed:
            print("   ✅ Version comparison works correctly")
    except Exception as e:
        print(f"   ❌ Version comparison test failed: {e}")
    
    # Test 5: Update checking (if internet available)
    print("5. Testing update checking...")
    try:
        if updater.check_internet_connection():
            update_info = updater.check_for_updates(show_no_updates=False)
            if update_info:
                print(f"   ✅ Update available: {update_info['version']}")
            else:
                print("   ✅ No updates available (current version is latest)")
        else:
            print("   ⚠️  Skipped (no internet connection)")
    except Exception as e:
        print(f"   ❌ Update check failed: {e}")
    
    print("\n🎉 Integration tests completed!")
    return True

def main():
    """Main test runner"""
    print("Energy Adjustment Calculator - Auto-Update System Tests")
    print("=" * 60)
    
    if not UPDATER_AVAILABLE:
        print("❌ Auto-updater system not available. Please check imports.")
        return
    
    # Run unit tests
    print("\n📋 Running Unit Tests...")
    unittest.main(argv=[''], exit=False, verbosity=2)
    
    # Run integration tests
    print("\n🔗 Running Integration Tests...")
    run_integration_tests()

if __name__ == "__main__":
    main()