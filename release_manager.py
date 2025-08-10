#!/usr/bin/env python3
"""
Release Manager for Energy Adjustment Calculator
Handles version management, changelog generation, and release preparation
"""

import os
import sys
import json
import subprocess
import zipfile
import shutil
from pathlib import Path
from datetime import datetime
import re

class ReleaseManager:
    def __init__(self):
        self.project_root = Path(__file__).parent
        self.version_file = self.project_root / "version.json"
        self.changelog_file = self.project_root / "CHANGELOG.md"
        
    def load_version_info(self):
        """Load current version information"""
        if self.version_file.exists():
            with open(self.version_file, 'r') as f:
                return json.load(f)
        return {
            "version": "1.0.0",
            "release_date": datetime.now().strftime("%Y-%m-%d"),
            "build_number": "001"
        }
    
    def save_version_info(self, version_data):
        """Save version information"""
        with open(self.version_file, 'w') as f:
            json.dump(version_data, f, indent=2)
    
    def increment_version(self, version_type="patch"):
        """Increment version number"""
        version_data = self.load_version_info()
        current_version = version_data["version"]
        
        # Parse version (major.minor.patch)
        major, minor, patch = map(int, current_version.split('.'))
        
        if version_type == "major":
            major += 1
            minor = 0
            patch = 0
        elif version_type == "minor":
            minor += 1
            patch = 0
        else:  # patch
            patch += 1
        
        new_version = f"{major}.{minor}.{patch}"
        
        # Update version data
        version_data["version"] = new_version
        version_data["release_date"] = datetime.now().strftime("%Y-%m-%d")
        version_data["build_number"] = str(int(version_data.get("build_number", "0")) + 1).zfill(3)
        
        return new_version, version_data
    
    def get_git_changes(self, since_tag=None):
        """Get git changes since last tag or commit"""
        try:
            if since_tag:
                cmd = ["git", "log", f"{since_tag}..HEAD", "--oneline"]
            else:
                cmd = ["git", "log", "--oneline", "-10"]  # Last 10 commits
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout.strip().split('\n')
            return []
        except Exception:
            return []
    
    def generate_changelog_entry(self, version, changes):
        """Generate changelog entry for version"""
        date_str = datetime.now().strftime("%Y-%m-%d")
        
        entry = f"\n## [{version}] - {date_str}\n\n"
        
        # Categorize changes
        features = []
        fixes = []
        improvements = []
        other = []
        
        for change in changes:
            change_lower = change.lower()
            if any(keyword in change_lower for keyword in ['feat:', 'feature:', 'add:', 'new:']):
                features.append(change)
            elif any(keyword in change_lower for keyword in ['fix:', 'bug:', 'patch:']):
                fixes.append(change)
            elif any(keyword in change_lower for keyword in ['improve:', 'enhance:', 'update:', 'refactor:']):
                improvements.append(change)
            else:
                other.append(change)
        
        if features:
            entry += "### ✨ New Features\n"
            for feature in features:
                entry += f"- {feature}\n"
            entry += "\n"
        
        if improvements:
            entry += "### 🚀 Improvements\n"
            for improvement in improvements:
                entry += f"- {improvement}\n"
            entry += "\n"
        
        if fixes:
            entry += "### 🐛 Bug Fixes\n"
            for fix in fixes:
                entry += f"- {fix}\n"
            entry += "\n"
        
        if other:
            entry += "### 📝 Other Changes\n"
            for change in other:
                entry += f"- {change}\n"
            entry += "\n"
        
        return entry
    
    def update_changelog(self, version, changes):
        """Update changelog file"""
        changelog_entry = self.generate_changelog_entry(version, changes)
        
        if self.changelog_file.exists():
            with open(self.changelog_file, 'r') as f:
                existing_content = f.read()
            
            # Insert new entry after the header
            lines = existing_content.split('\n')
            header_end = 0
            for i, line in enumerate(lines):
                if line.startswith('## [') or line.startswith('# Changelog'):
                    if line.startswith('# Changelog'):
                        header_end = i + 1
                    else:
                        header_end = i
                    break
            
            new_content = '\n'.join(lines[:header_end]) + changelog_entry + '\n'.join(lines[header_end:])
        else:
            new_content = f"# Changelog\n\nAll notable changes to the Energy Adjustment Calculator will be documented in this file.\n{changelog_entry}"
        
        with open(self.changelog_file, 'w') as f:
            f.write(new_content)
    
    def create_release_package(self, version):
        """Create release package"""
        release_dir = self.project_root / "releases" / f"v{version}"
        release_dir.mkdir(parents=True, exist_ok=True)
        
        # Files to include in release
        release_files = [
            'streamlit_app.py',
            'app.py',
            'launcher.py',
            'auto_updater.py',
            'version.json',
            'requirements.txt',
            'requirements_windows.txt',
            'templates/',
            '.streamlit/',
            'BUILD_INSTRUCTIONS.md',
            'WINDOWS_APP_GUIDE.md',
            'AUTO_UPDATE_GUIDE.md',
            'CHANGELOG.md'
        ]
        
        # Create release ZIP
        release_zip = release_dir / f"EnergyCalculator_v{version}.zip"
        
        with zipfile.ZipFile(release_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_pattern in release_files:
                file_path = self.project_root / file_pattern
                if file_path.exists():
                    if file_path.is_file():
                        zipf.write(file_path, file_pattern)
                    elif file_path.is_dir():
                        for sub_file in file_path.rglob('*'):
                            if sub_file.is_file():
                                arcname = str(sub_file.relative_to(self.project_root))
                                zipf.write(sub_file, arcname)
        
        return release_zip
    
    def create_update_package(self, version):
        """Create update package (smaller, only changed files)"""
        update_dir = self.project_root / "updates" / f"v{version}"
        update_dir.mkdir(parents=True, exist_ok=True)
        
        # Core files that typically get updated
        update_files = [
            'streamlit_app.py',
            'app.py',
            'launcher.py',
            'auto_updater.py',
            'version.json',
            'requirements.txt',
            'requirements_windows.txt',
            'templates/index.html',
            '.streamlit/config.toml'
        ]
        
        # Create update ZIP
        update_zip = update_dir / f"EnergyCalculator_Update_v{version}.zip"
        
        with zipfile.ZipFile(update_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_pattern in update_files:
                file_path = self.project_root / file_pattern
                if file_path.exists() and file_path.is_file():
                    zipf.write(file_path, file_pattern)
        
        return update_zip
    
    def prepare_release(self, version_type="patch", message=None):
        """Prepare a new release"""
        print(f"🚀 Preparing {version_type} release...")
        
        # Increment version
        new_version, version_data = self.increment_version(version_type)
        print(f"📦 New version: {new_version}")
        
        # Get changes
        changes = self.get_git_changes()
        if message:
            changes.insert(0, message)
        
        # Update version file
        self.save_version_info(version_data)
        print(f"✅ Updated version.json")
        
        # Update changelog
        self.update_changelog(new_version, changes)
        print(f"✅ Updated CHANGELOG.md")
        
        # Create release package
        release_zip = self.create_release_package(new_version)
        print(f"✅ Created release package: {release_zip}")
        
        # Create update package
        update_zip = self.create_update_package(new_version)
        print(f"✅ Created update package: {update_zip}")
        
        # Update installers
        self.update_installers()
        print(f"✅ Updated installer files")
        
        print(f"\n🎉 Release v{new_version} prepared successfully!")
        print(f"📁 Release files:")
        print(f"   - Full release: {release_zip}")
        print(f"   - Update package: {update_zip}")
        print(f"   - Updated installers in project root")
        
        return new_version, version_data
    
    def update_installers(self):
        """Update installer files with latest version"""
        try:
            # Regenerate ultimate installer
            subprocess.run([sys.executable, "create_ultimate_installer.py"], 
                         capture_output=True, check=True)
            
            # Regenerate bundle
            subprocess.run([sys.executable, "create_windows_bundle.py"], 
                         input="3\n", text=True, capture_output=True, check=True)
            
        except subprocess.CalledProcessError as e:
            print(f"⚠️  Warning: Failed to update installers: {e}")
    
    def list_releases(self):
        """List all available releases"""
        releases_dir = self.project_root / "releases"
        if not releases_dir.exists():
            print("No releases found.")
            return
        
        print("📦 Available Releases:")
        for release_dir in sorted(releases_dir.iterdir(), reverse=True):
            if release_dir.is_dir():
                version = release_dir.name
                release_files = list(release_dir.glob("*.zip"))
                if release_files:
                    file_size = release_files[0].stat().st_size / 1024 / 1024
                    print(f"   {version} ({file_size:.1f} MB)")
    
    def show_current_version(self):
        """Show current version information"""
        version_data = self.load_version_info()
        print(f"📋 Current Version Information:")
        print(f"   Version: {version_data['version']}")
        print(f"   Release Date: {version_data['release_date']}")
        print(f"   Build Number: {version_data['build_number']}")

def main():
    """Main CLI interface"""
    manager = ReleaseManager()
    
    if len(sys.argv) < 2:
        print("Energy Adjustment Calculator - Release Manager")
        print("=" * 50)
        print("Usage:")
        print("  python release_manager.py <command> [options]")
        print("")
        print("Commands:")
        print("  version                 - Show current version")
        print("  releases               - List all releases")
        print("  prepare <type> [msg]   - Prepare new release (patch/minor/major)")
        print("  changelog              - Generate changelog from git")
        print("")
        print("Examples:")
        print("  python release_manager.py version")
        print("  python release_manager.py prepare patch 'Bug fixes and improvements'")
        print("  python release_manager.py prepare minor 'New auto-update feature'")
        return
    
    command = sys.argv[1]
    
    if command == "version":
        manager.show_current_version()
    
    elif command == "releases":
        manager.list_releases()
    
    elif command == "prepare":
        version_type = sys.argv[2] if len(sys.argv) > 2 else "patch"
        message = sys.argv[3] if len(sys.argv) > 3 else None
        
        if version_type not in ["patch", "minor", "major"]:
            print("❌ Invalid version type. Use: patch, minor, or major")
            return
        
        manager.prepare_release(version_type, message)
    
    elif command == "changelog":
        changes = manager.get_git_changes()
        if changes:
            print("📝 Recent Changes:")
            for change in changes:
                print(f"   - {change}")
        else:
            print("No recent changes found.")
    
    else:
        print(f"❌ Unknown command: {command}")

if __name__ == "__main__":
    main()