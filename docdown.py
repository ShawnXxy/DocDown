#!/usr/bin/env python3

import argparse
import os
from pathlib import Path
from docx import Document
import shutil
import re
import logging
from datetime import datetime
from docx.oxml.shared import qn
from docx.oxml import OxmlElement

class ConversionStats:
    def __init__(self):
        self.total_files = 0
        self.successful_files = []
        self.failed_files = []
        self.total_images = 0
        self.failed_images = []
        self.file_image_counts = {}  # Track images per file

def setup_logging(log_dir):
    """Setup logging to both console and file"""
    os.makedirs(log_dir, exist_ok=True)
    
    # Create a timestamp for the log file
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'docdown_{timestamp}.log')
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    return logging.getLogger(__name__)

def get_heading_level(paragraph):
    """Determine the correct heading level based on the Word document structure"""
    try:
        style_name = paragraph.style.name
        if style_name.startswith('Heading'):
            return int(style_name.split()[-1])
        elif style_name == 'Title':
            return 1
        
        if hasattr(paragraph, '_element'):
            properties = paragraph._element.get_or_add_pPr()
            if properties is not None:
                run_props = None
                for run in paragraph.runs:
                    if hasattr(run, '_element') and run._element.rPr is not None:
                        run_props = run._element.rPr
                        break
                
                if run_props is not None:
                    # Check font size if available
                    sz_elem = run_props.find('.//w:sz', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    if sz_elem is not None:
                        size = int(sz_elem.get(qn('w:val'))) // 2
                        if size >= 20:
                            return 1
                        elif size >= 16:
                            return 2
                        elif size >= 14:
                            return 3
                
                    # Check if it's bold and might be a heading
                    b_elem = run_props.find('.//w:b', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    if b_elem is not None:
                        return 3
        
        return 0
    except Exception as e:
        logging.warning(f"Error determining heading level: {str(e)}")
        return 0

def extract_images(doc, image_dir, doc_name, logger, stats):
    """Extract images from the document and save them to image directory"""
    image_refs = {}
    file_image_count = 0  # Counter for this specific file
    
    try:
        for rel in doc.part.rels.values():
            # Only process image relationships and verify the target exists
            if ("image" in rel.reltype and 
                hasattr(rel, 'target_part') and 
                hasattr(rel.target_part, 'blob')):
                try:
                    # Get image extension
                    image_data = rel.target_part.blob
                    if not image_data:  # Skip if no actual image data
                        continue
                        
                    image_ext = rel.target_ref.split('.')[-1].lower()
                    
                    # Validate image extension
                    if image_ext not in ['png', 'jpg', 'jpeg', 'gif', 'bmp']:
                        image_ext = 'png'  # Default to PNG for unknown types
                    
                    file_image_count += 1  # Increment counter for this file
                    image_filename = f"{doc_name}_image_{file_image_count}.{image_ext}"
                    
                    # Save image
                    image_path = os.path.join(image_dir, image_filename)
                    with open(image_path, 'wb') as img_file:
                        img_file.write(image_data)
                    
                    image_refs[rel.rId] = f"./images/{image_filename}"
                    logger.debug(f"Saved image: {image_filename}")
                
                except Exception as e:
                    error_msg = f"Failed to extract image {file_image_count} from {doc_name}: {str(e)}"
                    logger.error(error_msg)
                    stats.failed_images.append((doc_name, error_msg))
        
        # Update statistics
        stats.file_image_counts[doc_name] = file_image_count
        stats.total_images += file_image_count
        
        if file_image_count > 0:
            logger.info(f"Extracted {file_image_count} images from {doc_name}")
        
        return image_refs
    
    except Exception as e:
        error_msg = f"Failed to process images in {doc_name}: {str(e)}"
        logger.error(error_msg)
        stats.failed_images.append((doc_name, error_msg))
        return image_refs

def convert_to_markdown(doc_path, output_dir, logger, stats):
    """Convert a single Word document to Markdown"""
    try:
        logger.info(f"Converting {doc_path} to markdown")
        
        # Validate file exists and is readable
        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"File not found: {doc_path}")
            
        # Check file size
        file_size = os.path.getsize(doc_path)
        if file_size == 0:
            raise ValueError(f"File is empty: {doc_path}")
            
        try:
            doc = Document(doc_path)
        except Exception as e:
            if "Package not found" in str(e):
                raise ValueError(
                    f"File appears to be corrupted or not a valid Word document: {doc_path}\n"
                    f"Please ensure the file is a valid .docx file and not password protected."
                )
            raise
            
        doc_name = Path(doc_path).stem
        
        # Create relative output directory structure
        rel_path = os.path.relpath(os.path.dirname(doc_path), start=os.path.dirname(output_dir))
        target_dir = os.path.join(output_dir, rel_path)
        os.makedirs(target_dir, exist_ok=True)
        
        # Create images directory if it doesn't exist
        image_dir = os.path.join(target_dir, "images")
        os.makedirs(image_dir, exist_ok=True)
        
        # Extract images
        image_refs = extract_images(doc, image_dir, doc_name, logger, stats)
        
        # Create markdown content
        md_content = []
        in_code_block = False
        
        for paragraph in doc.paragraphs:
            # Collect paragraph content in order
            paragraph_content = []
            
            for run in paragraph.runs:
                has_image = False
                
                # Check for images in this run
                if hasattr(run, '_element'):
                    for element in run._element.findall('.//w:drawing', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        try:
                            inline_or_anchor = element.find('.//wp:inline', {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'})
                            if inline_or_anchor is None:
                                inline_or_anchor = element.find('.//wp:anchor', {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'})
                            
                            if inline_or_anchor is not None:
                                blip = inline_or_anchor.find('.//a:blip', {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
                                if blip is not None:
                                    rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                    if rId in image_refs:
                                        docPr = inline_or_anchor.find('.//wp:docPr', {'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'})
                                        alt_text = docPr.get('descr', 'Image') if docPr is not None else 'Image'
                                        alt_text = alt_text.replace('\n', ' ').strip()
                                        
                                        # Store the image markdown
                                        paragraph_content.append({
                                            'type': 'image',
                                            'content': f"![{alt_text}]({image_refs[rId]})"
                                        })
                                        has_image = True
                        except Exception as e:
                            logger.warning(f"Failed to process inline image in {doc_name}: {str(e)}")
                
                # If run has text and wasn't completely replaced by an image
                if run.text.strip() and not has_image:
                    paragraph_content.append({
                        'type': 'text',
                        'content': run.text
                    })
            
            # Process the collected content
            if paragraph_content:
                # Check if paragraph is a heading
                heading_level = get_heading_level(paragraph)
                if heading_level > 0:
                    if in_code_block:
                        md_content.append("```")
                        md_content.append("")
                        in_code_block = False
                    
                    # Combine all text content for heading
                    heading_text = ''.join(item['content'] for item in paragraph_content if item['type'] == 'text')
                    md_content.append(f"{'#' * heading_level} {heading_text}")
                else:
                    # Handle regular paragraph or code block
                    is_code = (
                        paragraph.style.name.lower().startswith('code') or
                        all(run.font.name in ['Consolas', 'Courier New'] for run in paragraph.runs if run.font)
                    )
                    
                    if is_code:
                        if not in_code_block:
                            md_content.append("")
                            md_content.append("```")
                            in_code_block = True
                        # Combine all text content for code
                        code_text = ''.join(item['content'] for item in paragraph_content if item['type'] == 'text')
                        md_content.append(code_text.replace('\t', '    '))
                    else:
                        if in_code_block:
                            md_content.append("```")
                            md_content.append("")
                            in_code_block = False
                        
                        # If paragraph contains only one image
                        if len(paragraph_content) == 1 and paragraph_content[0]['type'] == 'image':
                            md_content.append("")
                            md_content.append(paragraph_content[0]['content'])
                            md_content.append("")
                        else:
                            # Combine content preserving order
                            combined_content = ''.join(item['content'] for item in paragraph_content)
                            if combined_content.strip():
                                md_content.append(combined_content)
            else:
                # Handle empty lines
                if not in_code_block:
                    md_content.append("")
        
        # Close any open code block
        if in_code_block:
            md_content.append("```")
            md_content.append("")
        
        # Clean up multiple consecutive blank lines
        cleaned_content = []
        last_was_blank = False
        for line in md_content:
            if line.strip():
                cleaned_content.append(line)
                last_was_blank = False
            elif not last_was_blank:
                cleaned_content.append(line)
                last_was_blank = True
        
        # Save markdown file
        output_path = os.path.join(target_dir, f"{doc_name}.md")
        with open(output_path, 'w', encoding='utf-8') as md_file:
            md_file.write('\n'.join(cleaned_content))
        
        logger.info(f"Successfully converted {doc_path} to {output_path}")
        stats.successful_files.append(doc_path)
        return output_path
    
    except FileNotFoundError as e:
        error_msg = f"File not found: {doc_path}"
        logger.error(error_msg)
        stats.failed_files.append((doc_path, error_msg))
        raise
        
    except ValueError as e:
        error_msg = str(e)
        logger.error(error_msg)
        stats.failed_files.append((doc_path, error_msg))
        raise
        
    except Exception as e:
        error_msg = f"Error converting {doc_path}: {str(e)}"
        logger.error(error_msg)
        stats.failed_files.append((doc_path, str(e)))
        raise

def process_directory(source_dir, output_dir, logger, stats):
    """Process directory recursively maintaining directory structure"""
    converted_files = []
    skipped_files = []
    
    logger.info(f"Processing directory: {source_dir}")
    for root, _, files in os.walk(source_dir):
        for file in files:
            if file.endswith(('.doc', '.docx')):
                stats.total_files += 1
                doc_path = os.path.join(root, file)
                try:
                    # Check if file is actually a Word document
                    with open(doc_path, 'rb') as f:
                        header = f.read(4)
                        # Check for ZIP header (valid .docx files are ZIP archives)
                        if header != b'PK\x03\x04':
                            error_msg = (
                                f"File has .docx extension but is not a valid Word document: {doc_path}\n"
                                "File may be corrupted or in an older .doc format."
                            )
                            logger.error(error_msg)
                            stats.failed_files.append((doc_path, error_msg))
                            skipped_files.append(doc_path)
                            continue
                    
                    output_path = convert_to_markdown(doc_path, output_dir, logger, stats)
                    converted_files.append((doc_path, output_path))
                    
                except Exception as e:
                    logger.error(f"Failed to convert {doc_path}: {str(e)}")
                    if doc_path not in skipped_files:
                        skipped_files.append(doc_path)
                    continue
    
    if skipped_files:
        logger.warning("\nSkipped files:")
        for file in skipped_files:
            logger.warning(f"- {file}")
    
    return converted_files

def print_summary(logger, stats):
    """Print detailed conversion summary"""
    # Add visual separator for better visibility
    logger.info("\n" + "="*80)
    logger.info("                               CONVERSION SUMMARY")
    logger.info("="*80 + "\n")
    
    # Main statistics with emoji indicators
    logger.info("📊 Overall Statistics:")
    logger.info(f"   • Total files processed: {stats.total_files}")
    logger.info(f"   • Successfully converted: {len(stats.successful_files)} ✅")
    logger.info(f"   • Failed conversions: {len(stats.failed_files)} {'❌' if stats.failed_files else '✅'}")
    
    # Print image statistics only if there were images
    if stats.total_images > 0:
        logger.info("\n📷 Image Statistics:")
        logger.info(f"   • Total images extracted: {stats.total_images}")
        logger.info("\n   Images per file:")
        for doc_name, count in stats.file_image_counts.items():
            if count > 0:  # Only show files that had images
                logger.info(f"   • {doc_name}: {count} images")
    
    if stats.successful_files:
        logger.info("\n✅ Successfully converted files:")
        for file in stats.successful_files:
            logger.info(f"   • {file}")
    
    if stats.failed_files:
        logger.info("\n❌ Failed conversions:")
        for file, error in stats.failed_files:
            logger.error(f"   • {file}")
            logger.error(f"     Error: {error}")
            
            # Add suggestions for common errors
            if "Package not found" in error or "not a valid Word document" in error:
                logger.info(f"\n   Suggestion for '{Path(file).name}':")
                logger.info("   • Check if the file is:")
                logger.info("     - A valid .docx file (not .doc)")
                logger.info("     - Not password protected")
                logger.info("     - Not corrupted")
    
    if stats.failed_images:
        logger.info("\n⚠️  Failed image extractions:")
        for doc, error in stats.failed_images:
            logger.error(f"   • {doc}")
            logger.error(f"     Error: {error}")
    
    # Add closing separator
    logger.info("\n" + "="*80)

def main():
    parser = argparse.ArgumentParser(description='Convert Word documents to Markdown')
    parser.add_argument('source', help='Source file or directory path')
    parser.add_argument('target', help='Target directory for output')
    parser.add_argument('--log-dir', default='logs', help='Directory for log files')
    
    args = parser.parse_args()
    
    # Setup logging
    logger = setup_logging(args.log_dir)
    stats = ConversionStats()
    
    try:
        # Create target directory if it doesn't exist
        os.makedirs(args.target, exist_ok=True)
        
        if os.path.isfile(args.source):
            # Process single file
            if args.source.endswith(('.doc', '.docx')):
                stats.total_files += 1
                output_path = convert_to_markdown(args.source, args.target, logger, stats)
                logger.info(f"Conversion complete: {args.source} -> {output_path}")
        else:
            # Process directory
            converted_files = process_directory(args.source, args.target, logger, stats)
        
        # Print final summary
        print_summary(logger, stats)
        
        # Return error code if any conversions failed
        return 1 if stats.failed_files or stats.failed_images else 0
    
    except Exception as e:
        logger.error(f"Fatal error: {str(e)}")
        return 1

if __name__ == '__main__':
    exit(main()) 