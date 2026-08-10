#!/usr/bin/env python3
"""
===============================================================================
EduScripts Utility: Directory File Flattener & Renamer (move_to_root_and_modify.py)
===============================================================================
Title       : Directory File Flattener & Renamer
Description : Recursively flattens subdirectories by moving all files into a
              root target directory, with optional batch file prefixing,
              automatic filename collision handling, and empty folder cleanup.
Author      : EduScripts Maintainer
Date        : 2026-08-08
Dependencies: Python 3.6+ (Standard Library: os, shutil, sys, argparse)
Usage       :
  Interactive Mode:
    python3 move_to_root_and_modify.py

  CLI Mode Examples:
    python3 move_to_root_and_modify.py /path/to/folder
    python3 move_to_root_and_modify.py /path/to/folder --prefix "2026_"
    python3 move_to_root_and_modify.py /path/to/folder --remove-empty-dirs --dry-run
    python3 move_to_root_and_modify.py /path/to/folder -p "Unit1_" -r -y
===============================================================================
"""

import os
import shutil
import sys
import argparse

def get_script_path():
    """Returns the absolute path of this script if available."""
    if '__file__' in globals() and __file__:
        return os.path.abspath(__file__)
    return None

def prompt_interactive_options(target_dir):
    """
    Launches an interactive guided wizard to collect options from the user
    when executed without command line flags.
    """
    print("=" * 65)
    print(" 🛠️  EduScripts Directory Flattener & Renamer Wizard")
    print("=" * 65)
    
    # 1. Target Directory
    cwd = os.getcwd()
    user_dir = input(f"\n📂 Target Directory Path [Press Enter for current directory: {cwd}]:\n > ").strip()
    if user_dir:
        target_dir = user_dir
    target_dir = os.path.abspath(target_dir)

    # 2. File Prefix
    print("\n🏷️  File Prefix Option:")
    print("   Would you like to prepend a prefix to all processed files (e.g. '2026_' or 'Unit1_')?")
    prefix = input("   Enter prefix string [Press Enter to skip]: ").strip()

    # 3. Remove Empty Directories
    print("\n🧹 Empty Directory Cleanup:")
    remove_dirs_input = input("   Automatically remove empty subdirectories after moving files? (y/N): ").strip().lower()
    remove_empty_dirs = remove_dirs_input in ('y', 'yes')

    # 4. Dry Run Mode
    print("\n🔍 Execution Mode:")
    dry_run_input = input("   Run in DRY-RUN mode (preview changes without modifying files)? (Y/n): ").strip().lower()
    dry_run = dry_run_input not in ('n', 'no')

    return target_dir, prefix, remove_empty_dirs, dry_run

def flatten_and_modify_directory(target_dir=None, prefix="", remove_empty_dirs=False, 
                                dry_run=False, assume_yes=False, include_hidden=False):
    """
    Main processing function to move nested files to root, apply prefixes, and cleanup empty dirs.
    """
    if target_dir is None:
        target_dir = os.getcwd()
    
    target_dir = os.path.abspath(target_dir)
    if not os.path.exists(target_dir):
        print(f"\n❌ Error: Target directory does not exist: {target_dir}")
        return False
    if not os.path.isdir(target_dir):
        print(f"\n❌ Error: Target path is not a directory: {target_dir}")
        return False

    script_path = get_script_path()

    print("\n" + "=" * 65)
    print(" ⚙️  Execution Parameters")
    print("=" * 65)
    print(f" Target Directory   : {target_dir}")
    print(f" File Prefix        : {repr(prefix) if prefix else '(None)'}")
    print(f" Remove Empty Dirs  : {'Yes' if remove_empty_dirs else 'No'}")
    print(f" Execution Mode     : {'DRY RUN (Preview Only)' if dry_run else 'REAL EXECUTION'}")
    print("=" * 65 + "\n")

    # Collect operations
    operations = []  # List of tuples: (action_type, src_path, dest_name, dest_path, rel_src)
    skipped_count = 0

    for current_dir, dirs, files in os.walk(target_dir, topdown=True):
        if not include_hidden:
            dirs[:] = [d for d in dirs if not d.startswith('.')]

        rel_dir = os.path.relpath(current_dir, target_dir)
        is_root = (rel_dir == '.')

        for file_name in files:
            # Skip hidden files
            if not include_hidden and file_name.startswith('.'):
                skipped_count += 1
                continue

            src_path = os.path.join(current_dir, file_name)

            # Skip moving/renaming the script itself
            if script_path and os.path.abspath(src_path) == script_path:
                continue

            # Determine desired target filename with prefix
            if prefix and not file_name.startswith(prefix):
                desired_name = prefix + file_name
            else:
                desired_name = file_name

            if is_root:
                # File is already in root folder
                if desired_name != file_name:
                    # Rename operation in root
                    operations.append(('RENAME_ROOT', src_path, file_name, desired_name, '.'))
            else:
                # File is in a subdirectory -> move to root
                operations.append(('MOVE_TO_ROOT', src_path, file_name, desired_name, rel_dir))

    if not operations and skipped_count == 0:
        print("ℹ️  No files found to process.")
        return True

    print(f"📋 Found {len(operations)} file operation(s) to process.")
    if skipped_count > 0:
        print(f"ℹ️  Skipping {skipped_count} hidden system file(s).")
    print()

    # Pre-calculate unique destination paths to resolve collisions
    execution_plan = []
    used_dest_names = set(os.listdir(target_dir))

    for action_type, src_path, orig_name, desired_name, rel_dir in operations:
        if action_type == 'RENAME_ROOT':
            # Remove original name temporarily from collision set since it's being renamed
            if orig_name in used_dest_names:
                used_dest_names.remove(orig_name)

        base_name, ext = os.path.splitext(desired_name)
        dest_name = desired_name
        counter = 1

        while dest_name in used_dest_names:
            dest_name = f"{base_name}_{counter}{ext}"
            counter += 1

        used_dest_names.add(dest_name)
        dest_path = os.path.join(target_dir, dest_name)
        execution_plan.append((action_type, src_path, orig_name, dest_name, dest_path, rel_dir))

    # Print action preview
    for action_type, src_path, orig_name, dest_name, dest_path, rel_dir in execution_plan:
        mode_str = "[WOULD PROCESS]" if dry_run else "[PROCESSING]"
        if action_type == 'MOVE_TO_ROOT':
            src_rel = os.path.join(rel_dir, orig_name)
            if orig_name != dest_name:
                print(f"  {mode_str} Move & Rename : {src_rel}  -->  {dest_name}")
            else:
                print(f"  {mode_str} Move to Root  : {src_rel}  -->  {dest_name}")
        elif action_type == 'RENAME_ROOT':
            print(f"  {mode_str} Rename in Root: {orig_name}  -->  {dest_name}")

    # Prompt confirmation if real execution and not assume_yes
    if not dry_run and not assume_yes:
        print()
        confirm = input("⚠️  Are you sure you want to execute these file changes? (y/N): ").strip().lower()
        if confirm not in ('y', 'yes'):
            print("\n❌ Operation cancelled by user.")
            return False

    print("\n🚀 Executing file operations...\n")
    success_count = 0

    # Execute file moves & renames
    for action_type, src_path, orig_name, dest_name, dest_path, rel_dir in execution_plan:
        if dry_run:
            success_count += 1
            continue

        try:
            if action_type == 'RENAME_ROOT':
                os.rename(src_path, dest_path)
                print(f"  ✅ Renamed : {orig_name}  -->  {dest_name}")
            else:
                shutil.move(src_path, dest_path)
                print(f"  ✅ Moved   : {os.path.join(rel_dir, orig_name)}  -->  {dest_name}")
            success_count += 1
        except Exception as e:
            print(f"  ❌ Failed  : {orig_name} - Error: {e}")

    # Optional empty directory cleanup
    cleaned_dir_count = 0
    if remove_empty_dirs:
        print("\n🧹 Checking for empty subdirectories to remove...")
        for current_dir, dirs, files in os.walk(target_dir, topdown=False):
            if not include_hidden:
                dirs[:] = [d for d in dirs if not d.startswith('.')]

            rel_dir = os.path.relpath(current_dir, target_dir)
            if rel_dir == '.':
                continue

            try:
                # Check if directory is empty (ignoring hidden files if include_hidden is False)
                remaining_items = os.listdir(current_dir)
                if not include_hidden:
                    remaining_items = [i for i in remaining_items if not i.startswith('.')]

                if not remaining_items:
                    if dry_run:
                        print(f"  [WOULD REMOVE DIR] {rel_dir}")
                        cleaned_dir_count += 1
                    else:
                        os.rmdir(current_dir)
                        print(f"  🗑️  Removed Empty Dir: {rel_dir}")
                        cleaned_dir_count += 1
            except Exception as e:
                print(f"  ⚠️ Could not remove folder {rel_dir}: {e}")

    # Final summary output
    print("\n" + "=" * 65)
    print(" ✨ Summary")
    print("=" * 65)
    action_verb = "Would process" if dry_run else "Processed"
    print(f" Total Files {action_verb}   : {success_count} / {len(execution_plan)}")
    if remove_empty_dirs:
        dir_verb = "Would remove" if dry_run else "Removed"
        print(f" Empty Folders {dir_verb} : {cleaned_dir_count}")
    if skipped_count > 0:
        print(f" System Files Skipped   : {skipped_count}")
    print("=" * 65 + "\n")

    return True

def main():
    parser = argparse.ArgumentParser(
        description="EduScripts Utility: Move nested subdirectory files to target root directory, apply optional prefixes, and clean up empty folders.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Interactive Guided Wizard:
    python3 move_to_root_and_modify.py

  CLI Direct Execution:
    python3 move_to_root_and_modify.py /path/to/dir --prefix "2026_" --remove-empty-dirs
    python3 move_to_root_and_modify.py /path/to/dir --dry-run
    python3 move_to_root_and_modify.py /path/to/dir -p "Unit1_" -r -y
"""
    )
    parser.add_argument("target_dir", nargs="?", default=None,
                        help="Path to target directory (default: current working directory)")
    parser.add_argument("-p", "--prefix", default="",
                        help="Prefix string to prepend to all processed files")
    parser.add_argument("-r", "--remove-empty-dirs", action="store_true",
                        help="Remove empty subdirectories after moving files")
    parser.add_argument("-n", "--dry-run", action="store_true",
                        help="Preview operations without modifying files or directories")
    parser.add_argument("-y", "--yes", action="store_true",
                        help="Skip interactive confirmation prompts")
    parser.add_argument("-a", "--include-hidden", action="store_true",
                        help="Include hidden files and subdirectories starting with '.'")
    parser.add_argument("-i", "--interactive", action="store_true",
                        help="Force interactive guided wizard mode")

    args = parser.parse_args()

    # Determine whether to run Interactive Wizard
    is_no_args = (len(sys.argv) == 1)
    if is_no_args or args.interactive:
        target_dir = args.target_dir or os.getcwd()
        target_dir, prefix, remove_empty_dirs, dry_run = prompt_interactive_options(target_dir)
        flatten_and_modify_directory(
            target_dir=target_dir,
            prefix=prefix,
            remove_empty_dirs=remove_empty_dirs,
            dry_run=dry_run,
            assume_yes=args.yes,
            include_hidden=args.include_hidden
        )
    else:
        flatten_and_modify_directory(
            target_dir=args.target_dir,
            prefix=args.prefix,
            remove_empty_dirs=args.remove_empty_dirs,
            dry_run=args.dry_run,
            assume_yes=args.yes,
            include_hidden=args.include_hidden
        )

if __name__ == '__main__':
    main()
