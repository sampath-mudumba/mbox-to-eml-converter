#!/usr/bin/env python3
"""
MBOX to EML Converter Script
===========================

This script converts .mbox files to individual .eml files.
Each email message in the mbox file is extracted and saved as a separate .eml file.

Features:
- Handles multiple emails in a single mbox file
- Preserves email headers and content
- Creates sanitized filenames for eml files
- Handles encoding issues gracefully
- Progress tracking for large mbox files
- Error handling for malformed messages

Usage:
    python mbox_to_eml.py input.mbox [output_directory]

Requirements:
    - Python 3.x (uses built-in mailbox module)

Author: AI Assistant
Date: August 2025
"""

import mailbox
import os
import sys
import re
from email import message_from_string
from email.generator import Generator
import argparse
from pathlib import Path

def sanitize_filename(filename, max_length=100):
    """
    Sanitize filename to be safe for filesystem.

    Args:
        filename (str): Original filename
        max_length (int): Maximum length for filename

    Returns:
        str: Sanitized filename
    """
    # Remove or replace invalid characters
    filename = re.sub(r'[<>:"/\|?*]', '_', filename)
    filename = re.sub(r'[\x00-\x1f]', '', filename)  # Remove control characters
    filename = filename.strip()

    # Truncate if too long
    if len(filename) > max_length:
        filename = filename[:max_length-4] + "..."

    return filename if filename else "unnamed_email"

def get_safe_filename(msg, index, output_dir):
    """
    Generate a safe filename for the email message.

    Args:
        msg: Email message object
        index (int): Message index
        output_dir (str): Output directory path

    Returns:
        str: Safe filename with full path
    """
    subject = msg.get('Subject', 'No Subject')
    sender = msg.get('From', 'Unknown Sender')

    # Extract email address from sender if present
    email_match = re.search(r'[\w\.-]+@[\w\.-]+', sender)
    sender_name = email_match.group(0) if email_match else sender

    # Create base filename from subject and sender
    base_name = f"{index:04d}_{sanitize_filename(subject)}_{sanitize_filename(sender_name)}"

    # Ensure unique filename
    filename = f"{base_name}.eml"
    filepath = os.path.join(output_dir, filename)

    counter = 1
    while os.path.exists(filepath):
        filename = f"{base_name}_{counter}.eml"
        filepath = os.path.join(output_dir, filename)
        counter += 1

    return filepath

def convert_mbox_to_eml(mbox_file, output_dir="eml_output", verbose=True):
    """
    Convert mbox file to individual eml files.

    Args:
        mbox_file (str): Path to the mbox file
        output_dir (str): Output directory for eml files
        verbose (bool): Enable verbose output

    Returns:
        tuple: (success_count, error_count, total_count)
    """
    if not os.path.exists(mbox_file):
        raise FileNotFoundError(f"MBOX file not found: {mbox_file}")

    # Create output directory if it doesn't exist
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    success_count = 0
    error_count = 0
    total_count = 0

    try:
        # Open the mbox file
        mbox = mailbox.mbox(mbox_file)

        if verbose:
            print(f"Processing mbox file: {mbox_file}")
            print(f"Output directory: {output_dir}")
            print("-" * 50)

        # Iterate through each message in the mbox
        for index, message in enumerate(mbox, 1):
            total_count = index

            try:
                # Generate safe filename
                eml_filepath = get_safe_filename(message, index, output_dir)

                # Write the message to eml file
                with open(eml_filepath, 'w', encoding='utf-8', errors='replace') as eml_file:
                    # Use Generator to properly format the email
                    gen = Generator(eml_file)
                    gen.flatten(message)

                success_count += 1

                if verbose:
                    subject = message.get('Subject', 'No Subject')[:50]
                    print(f"[{index:04d}] ✓ {os.path.basename(eml_filepath)}")

            except Exception as e:
                error_count += 1
                if verbose:
                    print(f"[{index:04d}] ✗ Error processing message: {str(e)}")
                continue

        if verbose:
            print("-" * 50)
            print(f"Conversion completed!")
            print(f"Total messages: {total_count}")
            print(f"Successfully converted: {success_count}")
            print(f"Errors: {error_count}")

    except Exception as e:
        raise RuntimeError(f"Error processing mbox file: {str(e)}")

    return success_count, error_count, total_count

def main():
    """
    Main function to handle command line arguments and execute conversion.
    """
    parser = argparse.ArgumentParser(
        description="Convert MBOX files to individual EML files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python mbox_to_eml.py emails.mbox
    python mbox_to_eml.py emails.mbox --output ./converted_emails
    python mbox_to_eml.py emails.mbox -o ./output --quiet
        """
    )

    parser.add_argument('mbox_file', 
                       help='Path to the MBOX file to convert')

    parser.add_argument('-o', '--output', 
                       default='eml_output',
                       help='Output directory for EML files (default: eml_output)')

    parser.add_argument('-q', '--quiet', 
                       action='store_true',
                       help='Suppress verbose output')

    parser.add_argument('--version', 
                       action='version', 
                       version='MBOX to EML Converter 1.0')

    args = parser.parse_args()

    try:
        success, errors, total = convert_mbox_to_eml(
            args.mbox_file, 
            args.output, 
            verbose=not args.quiet
        )

        # Exit with appropriate code
        if errors == 0:
            sys.exit(0)  # Success
        elif success > 0:
            sys.exit(1)  # Partial success
        else:
            sys.exit(2)  # Complete failure

    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(2)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(2)

if __name__ == "__main__":
    main()

# Alternative simple function for direct use
def simple_mbox_to_eml(mbox_path, output_dir="eml_files"):
    """
    Simple function to convert mbox to eml files.

    Args:
        mbox_path (str): Path to mbox file
        output_dir (str): Output directory
    """
    import mailbox
    import os
    from pathlib import Path

    # Create output directory
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Open mbox file
    mbox = mailbox.mbox(mbox_path)

    for i, message in enumerate(mbox, 1):
        # Create filename
        subject = message.get('Subject', 'No_Subject')
        # Clean subject for filename
        clean_subject = re.sub(r'[^a-zA-Z0-9_-]', '_', subject)[:30]
        filename = f"{i:04d}_{clean_subject}.eml"

        # Write eml file
        with open(os.path.join(output_dir, filename), 'w') as f:
            f.write(str(message))

        print(f"Converted: {filename}")

    print(f"Conversion complete. {i} messages converted.")
