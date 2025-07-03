import win32com.client
import os
import re
import json
import time
import psutil
import traceback
import pythoncom
from typing import Dict, List, Union, Any, Optional
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook


def check_file_in_use(file_path: str) -> bool:
    """Check if a file is being used"""
    if not os.path.exists(file_path):
        return False
    
    try:
        # Try to open the file in exclusive mode
        with open(file_path, 'r+b') as f:
            pass
        return False
    except (OSError, IOError):
        return True

def close_powerpoint_processes():
    """Close all PowerPoint processes"""
    print("🔄 Checking and closing PowerPoint processes...")
    closed_count = 0
    
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'].lower() in ['powerpnt.exe', 'powerpoint.exe']:
                print(f"   Found PowerPoint process PID: {proc.info['pid']}")
                proc.terminate()
                proc.wait(timeout=5)
                closed_count += 1
                print(f"   Closed PowerPoint process PID: {proc.info['pid']}")
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
            pass
    
    if closed_count > 0:
        print(f"✅ Closed {closed_count} PowerPoint processes")
        time.sleep(2)  # Wait for processes to fully close
    else:
        print("ℹ️  No PowerPoint processes found to close")

def generate_unique_filename(base_path: str) -> str:
    """Generate a unique filename"""
    if not os.path.exists(base_path):
        return base_path
    
    base_dir = os.path.dirname(base_path)
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    extension = os.path.splitext(base_path)[1]
    
    counter = 1
    while True:
        new_path = os.path.join(base_dir, f"{base_name}_{counter}{extension}")
        if not os.path.exists(new_path):
            return new_path
        counter += 1

def safe_save_presentation(pres, output_path: str, max_retries: int = 3) -> bool:
    """Safely save presentation with retry mechanism"""
    original_path = output_path
    
    for attempt in range(max_retries):
        try:
            print(f"💾 Attempting to save file (Attempt {attempt + 1}/{max_retries}): {output_path}")
            
            # Check if file is in use
            if check_file_in_use(output_path):
                print(f"⚠️  File is in use: {output_path}")
                
                if attempt == 0:
                    # First attempt: close PowerPoint processes
                    close_powerpoint_processes()
                    time.sleep(1)
                    continue
                else:
                    # Subsequent attempts: use new filename
                    output_path = generate_unique_filename(original_path)
                    print(f"🔄 Using new filename: {output_path}")
            
            # Try to save
            pres.SaveAs(output_path)
            print(f"✅ File saved successfully: {output_path}")
            return True
            
        except Exception as e:
            error_msg = str(e)
            print(f"❌ Save failed (Attempt {attempt + 1}): {error_msg}")
            
            if "being used" in error_msg.lower():
                if attempt < max_retries - 1:
                    # File in use error, try to resolve
                    print("🔄 File in use detected, attempting to resolve...")
                    close_powerpoint_processes()
                    
                    # Generate new filename
                    output_path = generate_unique_filename(original_path)
                    print(f"🔄 Using new filename: {output_path}")
                    time.sleep(2)
                    continue
            else:
                # Other errors, just retry
                if attempt < max_retries - 1:
                    print(f"⏳ Waiting {(attempt + 1) * 2} seconds before retry...")
                    time.sleep((attempt + 1) * 2)
                    continue
    
    print(f"💥 Save failed after {max_retries} attempts")
    return False


class PTMLParser:
    """PPT Template Markup Language (PTML) Parser"""
    
    # Marker type definitions - using ${} format, only keep direct replacement markers
    MARKERS = {
        'NO_CONVERT': r'\$\{([A-Za-z][A-Za-z0-9_]*)\}',  # ${Value} - No conversion
    }

def print_slide_content(slide, page_num: int):
    """Print all content of a slide"""
    print(f"\n{'='*80}")
    print(f"Page {page_num} Original Content:")
    print(f"{'='*80}")
    
    if slide.Shapes.Count == 0:
        print("  This page has no shapes")
        return
    
    for shape_idx, shape in enumerate(slide.Shapes, 1):
        print(f"\n📍 Shape {shape_idx}:")
        print(f"   Type: {shape.Type} ({get_shape_type_name(shape.Type)})")
        print(f"   Position: Left={shape.Left:.1f}, Top={shape.Top:.1f}")
        print(f"   Size: Width={shape.Width:.1f}, Height={shape.Height:.1f}")
        
        # Handle text frame content - improved version
        if shape.HasTextFrame:
            text_frame = shape.TextFrame
            if text_frame.HasText:
                # Get original text content without processing
                text_content = text_frame.TextRange.Text
                
                # Show complete original text (no truncation)
                print(f"   📝 Text Content (Length: {len(text_content)}):")
                print(f"      Original Text: {repr(text_content)}")
                print(f"      Display Text: 「{text_content}」")
                
                # If text is long, show by lines
                if len(text_content) > 100:
                    lines = text_content.split('\n')
                    print(f"      Line Display ({len(lines)} lines):")
                    for i, line in enumerate(lines, 1):
                        if line.strip():  # Only show non-empty lines
                            print(f"        Line {i}: 「{line}」")
                
                # Detailed marker detection
                print(f"   🔍 Marker Detection Results:")
                all_markers = []
                
                # Check each marker type
                import re
                for marker_type, pattern in PTMLParser.MARKERS.items():
                    try:
                        matches = re.findall(pattern, text_content)
                        if matches:
                            print(f"      ✅ {marker_type}: {matches}")
                            all_markers.extend([(marker_type, match) for match in matches])
                        else:
                            print(f"      ❌ {marker_type}: Not Found")
                    except Exception as e:
                        print(f"      ⚠️  {marker_type}: Detection Error - {e}")
                
                if not all_markers:
                    print(f"      ℹ️  No PTML markers found")
                
                # Additional check: find all possible $ markers
                simple_dollar_matches = re.findall(r'\$[^}]*\}?', text_content)
                if simple_dollar_matches:
                    print(f"   💡 All $ markers found: {simple_dollar_matches}")
                    
            else:
                print(f"   📝 Text Frame: Empty")
        
        # Handle table content - improved version
        if shape.HasTable:
            table = shape.Table
            print(f"   📊 Table: {table.Rows.Count} rows × {table.Columns.Count} columns")
            
            # Print complete table content, no truncation
            print(f"      Complete Table Content:")
            for row in range(1, table.Rows.Count + 1):
                row_data = []
                for col in range(1, table.Columns.Count + 1):
                    try:
                        cell = table.Cell(row, col)
                        if cell.Shape.HasTextFrame and cell.Shape.TextFrame.HasText:
                            cell_text = cell.Shape.TextFrame.TextRange.Text
                            # No truncation, show complete content and check markers
                            row_data.append(f"「{cell_text}」")
                            
                            # Check cell markers
                            cell_markers = []
                            for marker_type, pattern in PTMLParser.MARKERS.items():
                                matches = re.findall(pattern, cell_text)
                                if matches:
                                    cell_markers.extend([(marker_type, match) for match in matches])
                            
                            if cell_markers:
                                print(f"        Cell[{row},{col}] Original: {repr(cell_text)}")
                                print(f"        Cell[{row},{col}] Markers: {cell_markers}")
                                
                        else:
                            row_data.append("「Empty」")
                    except Exception as e:
                        row_data.append(f"「Error:{e}」")
                
                if row == 1:
                    print(f"        Header: {' | '.join(row_data)}")
                else:
                    print(f"        Row {row}: {' | '.join(row_data)}")
        
        # Handle chart content - improved version
        if shape.HasChart:
            chart = shape.Chart
            chart_type_name = get_chart_type_name(chart.ChartType)
            print(f"   📈 Chart: {chart_type_name} (Type Code: {chart.ChartType})")
            
            # Get chart title
            try:
                if chart.HasTitle:
                    title_text = chart.ChartTitle.Text
                    print(f"      Title Original: {repr(title_text)}")
                    print(f"      Title Display: 「{title_text}」")
                    
                    # Check title markers
                    title_markers = []
                    for marker_type, pattern in PTMLParser.MARKERS.items():
                        matches = re.findall(pattern, title_text)
                        if matches:
                            title_markers.extend([(marker_type, match) for match in matches])
                    
                    if title_markers:
                        print(f"      Title Markers: {title_markers}")
                else:
                    print(f"      Title: None")
            except Exception as e:
                print(f"      Title: Read Error - {e}")
            
            # Get series information
            try:
                series_count = chart.SeriesCollection.Count
                print(f"      Number of Series: {series_count}")
                
                for i in range(1, series_count + 1):  # Show all series
                    try:
                        series = chart.SeriesCollection(i)
                        series_name = str(series.Name)
                        print(f"      Series {i} Original: {repr(series_name)}")
                        print(f"      Series {i} Display: 「{series_name}」")
                        
                        # Check series name markers
                        series_markers = []
                        for marker_type, pattern in PTMLParser.MARKERS.items():
                            matches = re.findall(pattern, series_name)
                            if matches:
                                series_markers.extend([(marker_type, match) for match in matches])
                        
                        if series_markers:
                            print(f"      Series {i} Markers: {series_markers}")
                            
                    except Exception as e:
                        print(f"      Series {i}: Read Error - {e}")
                        
            except Exception as e:
                print(f"      Series Data: Read Error - {e}")
        
        # Handle image content - improved version
        if shape.Type == 13:  # Image type
            print(f"   🖼️  Image:")
            try:
                if hasattr(shape, 'Name'):
                    shape_name = shape.Name
                    print(f"      Name: 「{shape_name}」")
                    
                if hasattr(shape, 'AlternativeText'):
                    alt_text = shape.AlternativeText
                    if alt_text:
                        print(f"      Alt Text Original: {repr(alt_text)}")
                        print(f"      Alt Text Display: 「{alt_text}」")
                        
                        # Check alt text markers
                        alt_markers = []
                        for marker_type, pattern in PTMLParser.MARKERS.items():
                            matches = re.findall(pattern, alt_text)
                            if matches:
                                alt_markers.extend([(marker_type, match) for match in matches])
                        
                        if alt_markers:
                            print(f"      Alt Text Markers: {alt_markers}")
                    else:
                        print(f"      Alt Text: None")
                        
            except Exception as e:
                print(f"      Image Info: Read Error - {e}")
        
        # Handle other shapes - improved version
        if not shape.HasTextFrame and not shape.HasTable and not shape.HasChart and shape.Type != 13:
            try:
                shape_name = getattr(shape, 'Name', 'Unknown')
                print(f"   🔹 Other Shape: 「{shape_name}」")
                
                # Try to get other possible text attributes
                if hasattr(shape, 'TextFrame2'):
                    try:
                        if shape.TextFrame2.HasText:
                            text_content = shape.TextFrame2.TextRange.Text
                            print(f"      TextFrame2 Content Original: {repr(text_content)}")
                            print(f"      TextFrame2 Content Display: 「{text_content}」")
                            
                            # Check TextFrame2 markers
                            tf2_markers = []
                            for marker_type, pattern in PTMLParser.MARKERS.items():
                                matches = re.findall(pattern, text_content)
                                if matches:
                                    tf2_markers.extend([(marker_type, match) for match in matches])
                            
                            if tf2_markers:
                                print(f"      TextFrame2 Markers: {tf2_markers}")
                    except:
                        pass
                        
            except Exception as e:
                print(f"   🔹 Other Shape: Info Read Error - {e}")
    
    print(f"\n{'='*80}")
    print(f"Page {page_num} Content Scan Complete - Total {slide.Shapes.Count} Shapes")
    print(f"{'='*80}")

def get_shape_type_name(shape_type: int) -> str:
    """Get shape type name"""
    shape_types = {
        1: "AutoShape",
        2: "Callout",
        3: "Chart",
        4: "Comment",
        5: "Freeform",
        6: "Group",
        7: "Embedded OLE Object",
        8: "Form Control",
        9: "Line",
        10: "Linked OLE Object",
        11: "Linked Picture",
        12: "Media",
        13: "Picture",
        14: "Placeholder",
        15: "Polygon",
        16: "Polyline",
        17: "Text Box",
        18: "Table",
        19: "Text Effect"
    }
    return shape_types.get(shape_type, f"Unknown Type({shape_type})")

def get_chart_type_name(chart_type: int) -> str:
    """Get chart type name"""
    chart_types = {
        4: "Line Chart",
        5: "Pie Chart",
        51: "Column Chart",
        52: "Stacked Column Chart",
        53: "100% Stacked Column Chart",
        57: "Bar Chart",
        65: "Area Chart",
        68: "Scatter Chart",
        69: "Bubble Chart",
        70: "Doughnut Chart",
        72: "Radar Chart"
    }
    return chart_types.get(chart_type, f"Unknown Chart({chart_type})")

def process_ptml_template(ppt_path: str, template_data: Dict[str, Any], output_path: Optional[str] = None, page_numbers: Optional[List[int]] = None) -> bool:
    """Process PPT template, replace markers"""
    import win32com.client
    import os
    import time
    
    # Ensure input file exists
    if not os.path.exists(ppt_path):
        print(f"❌ Input file does not exist: {ppt_path}")
        return False
    
    # If output path is not specified, use input path
    if output_path is None:
        output_path = ppt_path
    
    # Ensure output directory exists
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"✅ Created output directory: {output_dir}")
        except Exception as e:
            print(f"❌ Failed to create output directory: {e}")
            return False
    
    print(f"🔄 Starting to process PPT template: {ppt_path}")
    
    try:
        # Create PowerPoint application instance
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = True  # Set to visible for debugging
        
        try:
            # Open source file
            print(f"📂 Opening file: {ppt_path}")
            source_pres = ppt.Presentations.Open(ppt_path)
            print(f"✅ File opened successfully")
            
            # Process all slides
            total_slides = source_pres.Slides.Count
            print(f"📊 Total pages: {total_slides}")
            
            # If page numbers are specified, process only those pages
            slides_to_process = page_numbers if page_numbers else range(1, total_slides + 1)
            
            for page_num in slides_to_process:
                if page_num > total_slides:
                    print(f"⚠️ Skipping invalid page number: {page_num} (exceeds total pages)")
                    continue
                
                print(f"\nProcessing page {page_num}:")
                slide = source_pres.Slides(page_num)
                
                # Print slide content (for debugging)
                print_slide_content(slide, page_num)
                
                # Process all shapes in the slide
                for shape_idx, shape in enumerate(slide.Shapes, 1):
                    print(f"\nProcessing shape {shape_idx}:")
                    
                    # Call corresponding processing function based on shape type
                    if shape.Type == 6:  # Group shape
                        process_group_shape(shape, template_data)
                    elif shape.HasTable:
                        process_table_shape(shape, template_data)
                    elif shape.HasChart:
                        process_chart_shape(shape, template_data)
                    elif shape.Type == 13:  # Image
                        process_image_shape(shape, template_data)
                    else:
                        process_text_shape(shape, template_data)
            
            # Save file
            print(f"\n💾 Saving file: {output_path}")
            if safe_save_presentation(source_pres, output_path):
                print(f"✅ File saved successfully: {output_path}")
                return True
            else:
                print(f"❌ File save failed")
                return False
                
        finally:
            try:
                # Close presentation
                source_pres.Close()
                print("✅ Closed presentation")
            except:
                pass
            
            try:
                # Exit PowerPoint
                ppt.Quit()
                print("✅ Exited PowerPoint")
            except:
                pass
            
            # Ensure all PowerPoint processes are closed
            print("🔄 Checking and closing PowerPoint processes...")
            close_powerpoint_processes()
            
    except Exception as e:
        print(f"Error creating PowerPoint application instance: {e}")
        print(f"Detailed error information: {traceback.format_exc()}")
        return False
    
    return True

def process_group_shape(shape, template_data: Dict[str, Any]):
    """Process markers in group shape (recursive processing of each sub-shape)"""
    if shape.Type != 6:  # Not a group shape
        return
    
    print(f"  📦 Processing group shape: {shape.Name}")
    
    try:
        # First check if the group shape itself has text
        if shape.HasTextFrame:
            print(f"    📝 Group shape itself has text frame")
            process_text_shape(shape, template_data)
        
        # Recursively process each sub-shape in the group
        for sub_shape_idx, sub_shape in enumerate(shape.GroupItems, 1):
            print(f"    🔧 Processing sub-shape {sub_shape_idx}: {get_shape_type_name(sub_shape.Type)}")
            
            # Recursively process nested group shapes
            if sub_shape.Type == 6:  # Nested group shape
                process_group_shape(sub_shape, template_data)
            
            # Process text of sub-shape
            if sub_shape.HasTextFrame:
                print(f"      📝 Sub-shape has text frame, processing...")
                process_text_shape(sub_shape, template_data)
            
            # Process table of sub-shape
            if sub_shape.HasTable:
                process_table_shape(sub_shape, template_data)
            
            # Process chart of sub-shape
            if sub_shape.HasChart:
                process_chart_shape(sub_shape, template_data)
            
            # Process image of sub-shape
            if sub_shape.Type == 13:  # Image shape
                process_image_shape(sub_shape, template_data)
                
    except Exception as e:
        print(f"    ⚠️   Error processing group shape: {e}")
        import traceback
        print(f"    Detailed error: {traceback.format_exc()}")

def get_case_insensitive_value(key: str, data_dict: Dict[str, Any], key_mapping: Dict[str, str]) -> Any:
    """Get case-insensitive key value"""
    # First try direct match
    if key in data_dict:
        return data_dict[key]
    
    # If direct match fails, try case-insensitive match
    if key.upper() in key_mapping:
        original_key = key_mapping[key.upper()]
        if original_key in data_dict:
            return data_dict[original_key]
    
    # If no match found, return None
    return None


def process_text_shape(shape, template_data: Dict[str, Any]):
    """Process text shape, directly replace marker value without any conversion"""
    if not shape.HasTextFrame:
        return
    
    text_frame = shape.TextFrame
    if not text_frame.HasText:
        return
    
    original_text = text_frame.TextRange.Text
    modified_text = original_text
    print(f"  🔍 Processing text: '{original_text}'")
    
    # Get key mapping dictionary
    key_mapping = template_data.get("_key_mapping", {})
    
    # Process all markers, directly replace without conversion
    no_convert_markers = re.findall(PTMLParser.MARKERS['NO_CONVERT'], modified_text)
    if no_convert_markers:
        print(f"   Found markers: {no_convert_markers}")
        for marker in no_convert_markers:
            value = get_case_insensitive_value(marker, template_data.get('TEXT', {}), key_mapping.get("TEXT", {}))
            if value is not None:
                modified_text = modified_text.replace(f"${{{marker}}}", str(value))
                print(f"     Direct replacement: ${{{marker}}} -> {value}")
            else:
                print(f"     No match found: ${{{marker}}}")
    
    # If text has changed, update
    if modified_text != original_text:
        text_frame.TextRange.Text = modified_text
        print(f"  ✅ Text updated")

def process_table_shape(shape, template_data: Dict[str, Any]):
    """Process markers in table shape without any conversion"""
    if not shape.HasTable:
        return
    
    try:
        table = shape.Table
        print(f"   Processing table: {table.Rows.Count} rows x {table.Columns.Count} columns")
        
        # Get key mapping dictionary
        key_mapping = template_data.get("_key_mapping", {})
        
        # Iterate through each cell in the table
        for row in range(1, table.Rows.Count + 1):
            for col in range(1, table.Columns.Count + 1):
                try:
                    cell = table.Cell(row, col)
                    if cell.Shape.HasTextFrame and cell.Shape.TextFrame.HasText:
                        cell_text = cell.Shape.TextFrame.TextRange.Text
                        original_text = cell_text
                        
                        # Process all markers, directly replace without conversion
                        no_convert_markers = re.findall(PTMLParser.MARKERS['NO_CONVERT'], cell_text)
                        if no_convert_markers:
                            print(f"     Found markers: {no_convert_markers}")
                            for marker in no_convert_markers:
                                value = get_case_insensitive_value(marker, template_data.get('TEXT', {}), key_mapping.get("TEXT", {}))
                                if value is not None:
                                    cell_text = cell_text.replace(f"${{{marker}}}", str(value))
                                    print(f"       Direct replacement: ${{{marker}}} -> {value}")
                                else:
                                    print(f"       No match found: ${{{marker}}}")
                            
                            # Only update if text has changed
                            if cell_text != original_text:
                                cell.Shape.TextFrame.TextRange.Text = cell_text
                                print(f"      ✅ Cell[{row},{col}] updated")
                except Exception as e:
                    print(f"     Error processing cell[{row},{col}]: {e}")
    except Exception as e:
        print(f"     Error processing table: {e}")
        print(f"     Detailed error information: {traceback.format_exc()}")

def process_chart_shape(shape, template_data: Dict[str, Any]):
    """Process markers in chart shape, including combo charts"""
    if not shape.HasChart:
        return
    
    chart = shape.Chart
    print(f"   Processing chart: {chart.ChartType}")
    
    # Get key mapping dictionary
    key_mapping = template_data.get("_key_mapping", {})
    
    try:
        # Process chart title
        try:
            if chart.HasTitle:
                title_text = chart.ChartTitle.Text
                original_text = title_text
                
                # Process all markers, directly replace without conversion
                no_convert_markers = re.findall(PTMLParser.MARKERS['NO_CONVERT'], title_text)
                if no_convert_markers:
                    print(f"     Found markers: {no_convert_markers}")
                    for marker in no_convert_markers:
                        value = get_case_insensitive_value(marker, template_data.get('TEXT', {}), key_mapping.get("TEXT", {}))
                        if value is not None:
                            title_text = title_text.replace(f"${{{marker}}}", str(value))
                            print(f"       Direct replacement: ${{{marker}}} -> {value}")
                        else:
                            print(f"       No match found: ${{{marker}}}")
                    
                    # Only update if text has changed
                    if title_text != original_text:
                        chart.ChartTitle.Text = title_text
                        print(f"      ✅ Chart title updated")
        except Exception as e:
            print(f"     Error processing chart title: {e}")
        
        # Process chart data update
        try:
            # Check if corresponding chart data exists
            chart_markers = re.findall(PTMLParser.MARKERS['NO_CONVERT'], str(shape.AlternativeText) if hasattr(shape, 'AlternativeText') else "")
            
            for marker in chart_markers:
                chart_data = get_case_insensitive_value(marker, template_data.get('CHARTS', {}), key_mapping.get("CHARTS", {}))
                if chart_data and isinstance(chart_data, dict):
                    print(f"     Updating chart data: {marker}")
                    update_chart_data(chart, chart_data)
                    
        except Exception as e:
            print(f"     Error processing chart data: {e}")
            
    except Exception as e:
        print(f"   Error processing chart: {e}")

def update_chart_data(chart, chart_data: Dict[str, Any]):
    """Update chart data, support combo charts"""
    try:
        chart_type = chart_data.get("type", "column")
        categories = chart_data.get("categories", [])
        
        print(f"    📊 Updating chart type: {chart_type}")
        print(f"    📊 Number of categories: {len(categories)}")
        
        # Update chart title
        if chart_data.get("title") and chart.HasTitle:
            chart.ChartTitle.Text = chart_data["title"]
            print(f"    ✅ Chart title updated to: {chart_data['title']}")
        
        # Process combo charts
        if chart_type.lower() == "combo":
            column_series = chart_data.get("column_series", [])
            line_series = chart_data.get("line_series", [])
            
            print(f"    📊 Column series: {len(column_series)}")
            print(f"    📊 Line series: {len(line_series)}")
            
            # Update data table (worksheet)
            try:
                chart_workbook = chart.ChartData.Workbook
                chart_worksheet = chart_workbook.Worksheets(1)
                
                # Clear existing data
                chart_worksheet.UsedRange.Clear()
                
                # Set categories (X-axis labels)
                for i, category in enumerate(categories, 2):  # Start from 2nd row
                    chart_worksheet.Cells(i, 1).Value = category
                
                # Set column chart data
                col_idx = 2  # Start from 2nd column
                for series_name in set(item["name"] for item in column_series):
                    chart_worksheet.Cells(1, col_idx).Value = series_name
                    
                    # Organize data by category
                    for i, category in enumerate(categories, 2):
                        # Find value corresponding to category
                        value = 0
                        for item in column_series:
                            if item["category"] == category and item["name"] == series_name:
                                value = item["value"]
                                break
                        chart_worksheet.Cells(i, col_idx).Value = value
                    
                    col_idx += 1
                
                # Set line chart data
                for series_name in set(item["name"] for item in line_series):
                    chart_worksheet.Cells(1, col_idx).Value = series_name
                    
                    # Organize data by category
                    for i, category in enumerate(categories, 2):
                        # Find value corresponding to category
                        value = 0
                        for item in line_series:
                            if item["category"] == category and item["name"] == series_name:
                                value = item["value"]
                                break
                        chart_worksheet.Cells(i, col_idx).Value = value
                    
                    col_idx += 1
                
                print(f"    ✅ Chart data updated")
                
                # Set chart series type
                try:
                    series_count = chart.SeriesCollection().Count
                    column_series_count = len(set(item["name"] for item in column_series))
                    
                    # Set first series as column chart
                    for i in range(1, min(column_series_count + 1, series_count + 1)):
                        series = chart.SeriesCollection(i)
                        series.ChartType = 51  # xlColumnClustered
                    
                    # Set subsequent series as line chart, using secondary axis
                    for i in range(column_series_count + 1, series_count + 1):
                        series = chart.SeriesCollection(i)
                        series.ChartType = 4   # xlLine
                        series.AxisGroup = 2   # Secondary axis
                    
                    print(f"    ✅ Chart series type set")
                    
                except Exception as e:
                    print(f"    ⚠️   Error setting chart series type: {e}")
                
            except Exception as e:
                print(f"    ⚠️   Error updating chart data: {e}")
        
        else:
            # Process other chart types
            print(f"    ℹ️   Unsupported chart type: {chart_type}")
            
    except Exception as e:
        print(f"    ❌ Error updating chart data: {e}")
        import traceback
        print(f"     Detailed error: {traceback.format_exc()}")

def process_image_shape(shape, template_data: Dict[str, Any]):
    """Process markers in image shape"""
    if shape.Type != 13:  # Not an image shape
        return
    
    # Get key mapping dictionary
    key_mapping = template_data.get("_key_mapping", {})
    
    # Check if alternative text contains markers
    try:
        if hasattr(shape, 'AlternativeText'):
            image_text = shape.AlternativeText
            image_markers = re.findall(PTMLParser.MARKERS['IMAGE'], image_text)
            if image_markers:
                print(f"   Found image markers: {image_markers}")
                for marker in image_markers:
                    new_image_path = get_case_insensitive_value(marker, template_data.get('IMAGES', {}), key_mapping.get("IMAGES", {}))
                    if new_image_path is not None and os.path.exists(new_image_path):
                        # Record original image position information
                        left, top = shape.Left, shape.Top
                        width, height = shape.Width, shape.Height
                        slide = shape.Parent
                        
                        # Delete original image
                        shape.Delete()
                        
                        # Add new image
                        slide.Shapes.AddPicture(
                            new_image_path,
                            False, True,
                            left, top, width, height
                        )
                        print(f"     Image replaced: {new_image_path}")
    except Exception as e:
        print(f"     Error processing image markers: {e}")

def read_excel_template(excel_path: str) -> Dict[str, Any]:
    """Read data from Excel template, keep original format (including percentage)"""
    try:
        print(f"\n📊 Reading Excel template: {excel_path}")
        
        # Use openpyxl directly to read Excel file
        wb = load_workbook(excel_path)
        
        template_data = {
            "TEXT": {},
            "DATES": {},
            "TABLES": {},
            "CHARTS": {},
            "IMAGES": {},
            "CONDITIONS": {}
        }
        
        # Create key mapping dictionary, for storing case mapping relationship
        key_mapping = {
            "TEXT": {},
            "DATES": {},
            "TABLES": {},
            "CHARTS": {},
            "IMAGES": {},
            "CONDITIONS": {}
        }
        
        # Process each sheet
        for sheet_name in wb.sheetnames:
            print(f"\n📑 Processing sheet: {sheet_name}")
            
            if sheet_name.lower() == "text":
                sheet = wb[sheet_name]
                
                # Find key and value column indices
                header_row = next(sheet.rows)
                key_col = None
                value_col = None
                for idx, cell in enumerate(header_row, 1):
                    if cell.value and str(cell.value).lower() == 'key':
                        key_col = idx
                    elif cell.value and str(cell.value).lower() == 'value':
                        value_col = idx
                
                if key_col is None or value_col is None:
                    print(f"  ⚠️  No necessary column found in sheet {sheet_name}")
                    continue
                
                # Process data from 2nd row (skip title row)
                for row in list(sheet.rows)[1:]:
                    key_cell = row[key_col - 1]
                    value_cell = row[value_col - 1]
                    
                    if not key_cell.value:
                        continue
                    
                    key = str(key_cell.value).strip()
                    
                    # Process value based on cell format
                    if value_cell.number_format and '%' in value_cell.number_format:
                        # If percentage format, directly use original string value
                        try:
                            # Get original value of cell
                            raw_value = value_cell.value
                            # If numeric type, convert to integer percentage format
                            if isinstance(raw_value, (int, float)):
                                # Multiply value by 100 and round to integer
                                percentage = round(raw_value * 100)
                                value = f"{percentage}%"
                            else:
                                # If not numeric, try to extract number from string
                                value = str(raw_value or '').strip()
                                if value:
                                    # If string contains number, try to convert to integer percentage
                                    try:
                                        # Remove all non-numeric characters (keep negative sign)
                                        num_str = ''.join(c for c in value if c.isdigit() or c == '-')
                                        if num_str:
                                            num_value = round(float(num_str))
                                            value = f"{num_value}%"
                                        elif not value.endswith('%'):
                                            value = f"{value}%"
                                    except ValueError:
                                        # If conversion fails, keep original value
                                        if not value.endswith('%'):
                                            value = f"{value}%"
                        except Exception as e:
                            print(f"  ⚠️   Error processing percentage value: {e}")
                            value = str(value_cell.value or '').strip()
                            if value and not value.endswith('%'):
                                value = f"{value}%"
                    else:
                        # Other cases use display value
                        value = str(value_cell.value or '').strip()
                    
                    # Save original key and uppercase key mapping relationship
                    key_mapping["TEXT"][key.upper()] = key
                    template_data["TEXT"][key] = value
                    print(f"  📝 Text: {key} -> {value}")
            
            elif sheet_name.lower() == "dates":
                # Process date data
                sheet = wb[sheet_name]
                for row in list(sheet.rows)[1:]:  # Skip title row
                    if not row[0].value:  # Check if key exists
                        continue
                    
                    key = str(row[0].value).strip()
                    value = row[1].value
                    
                    if isinstance(value, (datetime, str)):
                        if isinstance(value, datetime):
                            value = value.strftime('%Y-%m-%d')
                        else:
                            value = value.strip()
                        
                        key_mapping["DATES"][key.upper()] = key
                        template_data["DATES"][key] = value
                        print(f"  📅 Date: {key} -> {value}")
            
            elif sheet_name.lower() == "combo_charts":
                # Process combo chart data (column chart + line chart)
                sheet = wb[sheet_name]
                current_chart = None
                chart_data = {}
                
                for row in list(sheet.rows)[1:]:  # Skip title row
                    if not row[0].value:
                        continue
                    
                    chart_name = str(row[0].value).strip()
                    category = str(row[1].value or '').strip()
                    series_type = str(row[2].value or '').strip()
                    series_name = str(row[3].value or '').strip()
                    value = row[4].value
                    chart_type = str(row[5].value or '').strip()
                    title = str(row[6].value or '').strip()
                    
                    # Initialize chart data structure
                    if chart_name not in chart_data:
                        chart_data[chart_name] = {
                            "type": chart_type or "combo",
                            "title": title,
                            "categories": [],
                            "column_series": [],  # Column chart series
                            "line_series": []     # Line chart series
                        }
                    
                    # Add category
                    if category and category not in chart_data[chart_name]["categories"]:
                        chart_data[chart_name]["categories"].append(category)
                    
                    # Add series data
                    if series_type and series_name and value is not None:
                        series_data = {
                            "name": series_name,
                            "category": category,
                            "value": float(value) if isinstance(value, (int, float)) else 0
                        }
                        
                        if series_type.lower() == "column":
                            chart_data[chart_name]["column_series"].append(series_data)
                        elif series_type.lower() == "line":
                            chart_data[chart_name]["line_series"].append(series_data)
                
                # Save to template data
                for chart_name, data in chart_data.items():
                    key_mapping["CHARTS"][chart_name.upper()] = chart_name
                    template_data["CHARTS"][chart_name] = data
                    print(f"  📊 Combo chart: {chart_name}")
                    print(f"     Type: {data['type']}")
                    print(f"     Title: {data['title']}")
                    print(f"     Categories: {data['categories']}")
                    print(f"     Column series: {len(data['column_series'])}")
                    print(f"     Line series: {len(data['line_series'])}")
            
            elif sheet_name.lower() == "revenue_data":
                # Process revenue data table
                sheet = wb[sheet_name]
                revenue_table_data = {
                    "headers": [],
                    "data": []
                }
                
                # Read header
                header_row = next(sheet.rows)
                for cell in header_row:
                    if cell.value:
                        revenue_table_data["headers"].append(str(cell.value))
                
                # Read data rows
                for row in list(sheet.rows)[1:]:  # Skip title row
                    row_data = []
                    for cell in row:
                        if cell.value is not None:
                            row_data.append(str(cell.value))
                        else:
                            row_data.append("")
                    if any(row_data):  # If row has data
                        revenue_table_data["data"].append(row_data)
                
                # Save to template data
                key_mapping["TABLES"]["REVENUE_DATA"] = "revenue_data"
                template_data["TABLES"]["revenue_data"] = revenue_table_data
                print(f"  📊 Revenue data table: {len(revenue_table_data['data'])} rows of data")
                print(f"     Headers: {revenue_table_data['headers']}")
            
            elif sheet_name.lower() == "tables":
                # Process table data
                sheet = wb[sheet_name]
                current_table = None
                headers = []
                data = []
                
                for row in sheet.rows:
                    if not row[0].value:
                        continue
                    
                    first_cell = str(row[0].value).strip()
                    if first_cell.lower() == 'table_name':
                        if current_table and headers:
                            template_data["TABLES"][current_table] = {
                                "headers": headers,
                                "data": data
                            }
                        current_table = str(row[1].value).strip()
                        headers = []
                        data = []
                    elif first_cell.lower() == 'header':
                        headers = [str(cell.value).strip() for cell in row[1:] if cell.value]
                    else:
                        row_data = []
                        for cell in row:
                            if cell.value is None:
                                row_data.append("")
                            else:
                                row_data.append(str(cell.value).strip())
                        if any(row_data):
                            data.append(row_data)
                
                if current_table and headers:
                    template_data["TABLES"][current_table] = {
                        "headers": headers,
                        "data": data
                    }
            
            elif sheet_name.lower() == "images":
                # Process image path
                sheet = wb[sheet_name]
                for row in list(sheet.rows)[1:]:  # Skip title row
                    if not row[0].value or not row[1].value:
                        continue
                    
                    key = str(row[0].value).strip()
                    path = str(row[1].value).strip()
                    
                    key_mapping["IMAGES"][key.upper()] = key
                    template_data["IMAGES"][key] = path
                    print(f"  🖼️  Image: {key} -> {path}")
        
        # Add key mapping to template_data
        template_data["_key_mapping"] = key_mapping
        print("\n✅ Excel template data read completed")
        return template_data
        
    except Exception as e:
        print(f"\n❌ Error reading Excel template: {e}")
        import traceback
        print(f"Detailed error information: {traceback.format_exc()}")
        raise


if __name__ == "__main__":
    # File path
    ppt_file = r"D:\pythonProject\LanchainProject\tests\ppt_chuli\无仓年度PPT-模版.pptx"
    excel_template = r"D:\pythonProject\LanchainProject\tests\ppt_chuli\template_data.xlsx"  # Excel template path
    output_path = r"D:\pythonProject\LanchainProject\tests\ppt_chuli\生成的单页报告.pptx"
    
    # Read Excel template data
    template_data = read_excel_template(excel_template)
    
    # Process PPT
    process_ptml_template(ppt_file, template_data, output_path, page_numbers=[1,2,3,4,5,6])