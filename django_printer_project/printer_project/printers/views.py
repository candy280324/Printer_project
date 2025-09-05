import io
from datetime import datetime
from zoneinfo import ZoneInfo
import win32print
import win32ui
from PIL import Image, ImageDraw, ImageFont, ImageWin
from django.shortcuts import render
from django.http import HttpResponse
from django.views.decorators.csrf import csrf_exempt
from .forms import GatePassForm
from django.utils import timezone

LABELS = [
    "Gate Pass No",
    "Emp code",
    "Emp Name",
    "Department",
    "Date & Out Time",
    "Purpose of Going",
    "Remarks",
]
VISITORLABELS = [
    "Date",
    "Visitor Pass No",
    "Name of the Visitor",
    "Place / Mobile No",
    "Purpose of Visit",
    "Whom to Meet",
    "Department",
    "Belongings",
    "In Time",
]
UNIFORMLABELS = [
    "Date",
    "Emp code",
    "Emp Name",
    "Department",
    "Particulars",
    "Issued Qty",
    "From Date",
    "To Date"
]

# Printer and page settings
PRINTER_NAME = "POS-80C"  # Change this if your printer name is different
DPI = 203
PAGE_W = 812  # 4 inches * 203 DPI (Width of 4-inch paper in pixels)
PAGE_H = 1200  # Initial page height, can be adjusted dynamically based on content
MARGIN = 20  # Margin for top, bottom, left, and right
ROW_H = 40  # Reduced row height to fit more content (in pixels)
FONT_SIZE_BIG=35
FONT_SIZE_NORMAL = 30  # Font size for content (adjust as needed)
FONT_SIZE_SMALL = 23  # Smaller font for labels

# Load fonts
FONT_PATH = "arial.ttf"  # Full path to the font file
big_font = ImageFont.truetype("arialbd.ttf", FONT_SIZE_BIG)
normal_font = ImageFont.truetype(FONT_PATH, FONT_SIZE_NORMAL)
small_font = ImageFont.truetype(FONT_PATH, FONT_SIZE_SMALL)

@csrf_exempt
def gatepass(request):
    if request.method == "POST":
        values = [
            request.POST.get('index_no'),
            request.POST.get('qr_code_result'),
            request.POST.get('empname'),
            request.POST.get('department'),
            datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%d-%m-%Y %H:%M:%S'),
            # timezone.now().strftime('%d-%m-%Y %H:%M:%S'),
            request.POST.get("pupose_of_gng"),
            request.POST.get("remarks")
        ]
        try:
            if request.POST.get('index_no')>"0":
                _print_gatepass(labels=LABELS, values=values)
                return render(request, "printers/gatepass_form.html", {"status": "✅ Printed"})
        except Exception as e:
            return render(request, "printers/gatepass_form.html", {"error": f"Printer error: {e}"})
    else:
        return render(request, "printers/gatepass_form.html")
def visitorpass(request):
    if request.method == "POST":
        values = [
            timezone.now().strftime('%d-%m-%Y'),
            request.POST.get('index_no'),
            request.POST.get('visitor'),
            request.POST.get('mobile_no'),
            request.POST.get('purpose'),
            request.POST.get('empcodeName'),
            request.POST.get("department_name"),
            request.POST.get("belongings"),
            datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%H:%M'),
            # timezone.now().strftime('%H:%M'),
        ]
        try:
            if request.POST.get('index_no')>"0":
                _print_visitpass(labels=VISITORLABELS, values=values)
                return render(request, "printers/visitor_pass.html", {"status": "✅ Printed"})
            else:
                return render(request, "printers/visitor_pass.html", {"error": "Database Error occured"})
        except Exception as e:
            return render(request, "printers/visitor_pass.html", {"error": f"Printer error: {e}"})
    else:
        return render(request, "printers/visitor_pass.html")
def uniformslip(request):
    if request.method == "POST":
        from_date_str = request.POST.get('fromdate')
        to_date_str = request.POST.get('todate')
        from_date = datetime.strptime(from_date_str, "%Y-%m-%d") if from_date_str else None
        to_date = datetime.strptime(to_date_str, "%Y-%m-%d") if to_date_str else None
        rom_date_formatted = from_date.strftime("%d-%m-%Y") if from_date else None
        to_date_formatted = to_date.strftime("%d-%m-%Y") if to_date else None
        values = [
            timezone.now().strftime('%d-%m-%Y'),
            request.POST.get('qr_code_result'),
            request.POST.get('empname'),
            request.POST.get('department'),
            request.POST.get('particulars'),
            request.POST.get('issuedqty'),
            rom_date_formatted,
            to_date_formatted
        ]
        try:
            _print_uniformpass(labels=UNIFORMLABELS, values=values)
            return render(request, "printers/uniform_slip.html", {"status": "✅ Printed"})
        except Exception as e:
            return render(request, "printers/uniform_slip.html", {"error": f"Printer error: {e}"})
    else:
        return render(request, "printers/uniform_slip.html")

# Function to wrap text
def wrap_text(draw, text, font, max_chars_per_line):
    """Wrap text to avoid cutting words in the middle and ensure each line has no more than max_chars_per_line characters."""
    words = text.split()
    lines = []
    current_line = ''
    
    for word in words:
        # Check the length of the current line if we add this word
        test_line = f"{current_line} {word}".strip()
        
        # If the length of the test line is within the max character limit
        if len(test_line) <= max_chars_per_line:
            current_line = test_line
        else:
            # If the word doesn't fit, start a new line with the current word
            if current_line:
                lines.append(current_line)
            current_line = word

    # Add the last line if any
    if current_line:
        lines.append(current_line)
    
    return lines


# Function to print gatepass
def _print_gatepass(labels, values):
    img = Image.new("L", (PAGE_W, PAGE_H), 255)  # Create an image buffer for printing
    d = ImageDraw.Draw(img)

    y = MARGIN  # Start Y position with margin

    # Header Section
    d.text((MARGIN+50, y), "Medha Servo Drives Pvt. Ltd.", font=big_font, fill=0)
    y += FONT_SIZE_NORMAL + 20
    d.text((MARGIN+60, y), "Associate Movement Gate Pass", font=normal_font, fill=0)
    y += FONT_SIZE_NORMAL + 20  # Extra space after header

    # Left and Right Column X-positions
    left_col_x = MARGIN
    right_col_x = PAGE_W // 2 - 180  # Adjust right column to use the entire remaining width

    line_spacing = small_font.size + 2  # Tight line spacing

    row_height = 50  # Adjust row height as per your needs

    # Loop through each label and value
    for label, value in zip(labels, values):
        # Wrap value text to calculate required row height
        max_val_chars = 23
        wrapped_lines = wrap_text(d, str(value), small_font, max_val_chars)
        num_lines = len(wrapped_lines)
        row_height = 10 + num_lines * (small_font.size + 2) + 10  # padding top & bottom

        top_y = y
        bottom_y = y + row_height

        # Draw top border
        d.line([(left_col_x, top_y), (right_col_x + 345, top_y)], fill=0, width=1)

        # Draw vertical lines
        d.line([(left_col_x, top_y), (left_col_x, bottom_y)], fill=0, width=1)        # Left
        d.line([(right_col_x, top_y), (right_col_x, bottom_y)], fill=0, width=1)      # Middle
        right_x = right_col_x + 345
        d.line([(right_x, top_y), (right_x, bottom_y)], fill=0, width=1)              # Right

        # Draw label
        d.text((left_col_x + 10, top_y + 10), label, font=small_font, fill=0)

        # Draw wrapped value line by line
        for i, line in enumerate(wrapped_lines):
            line_y = top_y + 10 + i * (small_font.size + 2)
            d.text((right_col_x + 10, line_y), line, font=small_font, fill=0)

        # Draw bottom border
        d.line([(left_col_x, bottom_y), (right_x, bottom_y)], fill=0, width=1)

        y += row_height  # Move down by full row height
    # Footer Section
    footer_text = "Authorized by Security ASO"
    footer_width = d.textbbox((0, 0), footer_text, font=small_font)[2]
    d.text(((PAGE_W - footer_width) // 2, y + 20), footer_text, font=small_font, fill=0)

    # Crop and print
    img = img.crop((0, 0, PAGE_W, y + 60))

    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(PRINTER_NAME)
    hDC.StartDoc("GatePass")
    hDC.StartPage()
    dib = ImageWin.Dib(img)
    dib.draw(hDC.GetHandleOutput(), (0, 0, PAGE_W, y + 60))
    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()

    _send_cut_command()
def _print_visitpass(labels, values):
    img = Image.new("L", (PAGE_W, PAGE_H), 255)  # Create an image buffer for printing
    d = ImageDraw.Draw(img)

    y = MARGIN  # Start Y position with margin

    # Header Section
    d.text((MARGIN+50, y), "Medha Servo Drives Pvt. Ltd.", font=big_font, fill=0)
    # d.text((MARGIN, y), "Medha Servo Drives Pvt. Ltd.", font=normal_font, fill=0)
    y += FONT_SIZE_NORMAL + 20
    d.text((MARGIN+200, y), "Visitor Pass", font=normal_font, fill=0)
    y += FONT_SIZE_NORMAL + 20  # Extra space after header

    # Left and Right Column X-positions
    left_col_x = MARGIN
    right_col_x = PAGE_W // 2 - 180  # Adjust right column to use the entire remaining width

    line_spacing = small_font.size + 2  # Tight line spacing

    row_height = 50  # Adjust row height as per your needs

    # Loop through each label and value
    for label, value in zip(labels, values):
        # Wrap value text to calculate required row height
        max_val_chars = 23
        wrapped_lines = wrap_text(d, str(value), small_font, max_val_chars)
        num_lines = len(wrapped_lines)
        row_height = 10 + num_lines * (small_font.size + 2) + 10  # padding top & bottom

        top_y = y
        bottom_y = y + row_height

        # Draw top border
        d.line([(left_col_x, top_y), (right_col_x + 345, top_y)], fill=0, width=1)

        # Draw vertical lines
        d.line([(left_col_x, top_y), (left_col_x, bottom_y)], fill=0, width=1)        # Left
        d.line([(right_col_x, top_y), (right_col_x, bottom_y)], fill=0, width=1)      # Middle
        right_x = right_col_x + 345
        d.line([(right_x, top_y), (right_x, bottom_y)], fill=0, width=1)              # Right

        # Draw label
        d.text((left_col_x + 10, top_y + 10), label, font=small_font, fill=0)

        # Draw wrapped value line by line
        for i, line in enumerate(wrapped_lines):
            line_y = top_y + 10 + i * (small_font.size + 2)
            d.text((right_col_x + 10, line_y), line, font=small_font, fill=0)

        # Draw bottom border
        d.line([(left_col_x, bottom_y), (right_x, bottom_y)], fill=0, width=1)

        y += row_height  # Move down by full row height

    # Footer Section
    footer_text = "Sign. Of the Visitor             Sign. Of the Person Visited"
    footer_width = d.textbbox((0, 0), footer_text, font=small_font)[2]

    # Adjust the X position and move it left
    shift_left = 100
    x_position = ((PAGE_W - footer_width) // 2) - shift_left
    d.text((x_position, y + 10), footer_text, font=small_font, fill=0)

    # Second Footer Text (Move the Y position down to avoid overlap)
    footer_text_2 = "Sign. Of the Security"
    footer_width_2 = d.textbbox((0, 0), footer_text_2, font=small_font)[2]

    # Adjust the X position again (optional)
    x_position_2 = ((PAGE_W - footer_width_2) // 2) - shift_left

    # Move the Y position further down to avoid overlap with the first footer text
    space_between_texts = 20  # Add a fixed space between the two footer texts
    y_footer_2 = y + 10 + small_font.size + space_between_texts  # Increase y-position for second footer text

    # Draw the second footer text
    d.text((x_position_2, y_footer_2), footer_text_2, font=small_font, fill=0)

    # Ensure enough space for footer when cropping
    img = img.crop((0, 0, PAGE_W, y_footer_2 + small_font.size + 20))  # Adjust cropping to include footer

    # Print the image
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(PRINTER_NAME)
    hDC.StartDoc("GatePass")
    hDC.StartPage()
    dib = ImageWin.Dib(img)
    dib.draw(hDC.GetHandleOutput(), (0, 0, PAGE_W, y_footer_2 + small_font.size + 20))  # Adjust size to match new cropping
    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()

    _send_cut_command()
def _print_uniformpass(labels, values):
    img = Image.new("L", (PAGE_W, PAGE_H), 255)  # Create an image buffer for printing
    d = ImageDraw.Draw(img)

    y = MARGIN  # Start Y position with margin

    # Header Section
    d.text((MARGIN+50, y), "Medha Servo Drives Pvt. Ltd.", font=big_font, fill=0)
    # d.text((MARGIN, y), "Medha Servo Drives Pvt. Ltd.", font=normal_font, fill=0)
    y += FONT_SIZE_NORMAL + 20
    d.text((MARGIN, y), "Uniform Permission Slip", font=normal_font, fill=0)
    y += FONT_SIZE_NORMAL + 20  # Extra space after header

    # Left and Right Column X-positions
    left_col_x = MARGIN
    right_col_x = PAGE_W // 2 - 180  # Adjust right column to use the entire remaining width

    line_spacing = small_font.size + 2  # Tight line spacing

    row_height = 50  # Adjust row height as per your needs

    # Loop through each label and value
    for label, value in zip(labels, values):
        # Wrap value text to calculate required row height
        max_val_chars = 23
        wrapped_lines = wrap_text(d, str(value), small_font, max_val_chars)
        num_lines = len(wrapped_lines)
        row_height = 10 + num_lines * (small_font.size + 2) + 10  # padding top & bottom

        top_y = y
        bottom_y = y + row_height

        # Draw top border
        d.line([(left_col_x, top_y), (right_col_x + 345, top_y)], fill=0, width=1)

        # Draw vertical lines
        d.line([(left_col_x, top_y), (left_col_x, bottom_y)], fill=0, width=1)        # Left
        d.line([(right_col_x, top_y), (right_col_x, bottom_y)], fill=0, width=1)      # Middle
        right_x = right_col_x + 345
        d.line([(right_x, top_y), (right_x, bottom_y)], fill=0, width=1)              # Right

        # Draw label
        d.text((left_col_x + 10, top_y + 10), label, font=small_font, fill=0)

        # Draw wrapped value line by line
        for i, line in enumerate(wrapped_lines):
            line_y = top_y + 10 + i * (small_font.size + 2)
            d.text((right_col_x + 10, line_y), line, font=small_font, fill=0)

        # Draw bottom border
        d.line([(left_col_x, bottom_y), (right_x, bottom_y)], fill=0, width=1)

        y += row_height  # Move down by full row height
    # Footer Section
    footer_text = "Authorized by"
    footer_width = d.textbbox((0, 0), footer_text, font=small_font)[2]
    d.text(((PAGE_W - footer_width) // 2, y + 20), footer_text, font=small_font, fill=0)

    # Crop and print
    img = img.crop((0, 0, PAGE_W, y + 60))

    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(PRINTER_NAME)
    hDC.StartDoc("GatePass")
    hDC.StartPage()
    dib = ImageWin.Dib(img)
    dib.draw(hDC.GetHandleOutput(), (0, 0, PAGE_W, y + 60))
    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()

    _send_cut_command()



# Helper function to send cut command to the printer
def _send_cut_command():
    """Send cut command to the printer after printing."""
    feed_and_cut_cmd = b'\x1B\x64\x05' + b'\x1D\x56\x00'  # Feed 5 lines + Full Cut
    printer = win32print.OpenPrinter(PRINTER_NAME)
    hJob = win32print.StartDocPrinter(printer, 1, ("Cut", None, "RAW"))
    win32print.StartPagePrinter(printer)
    win32print.WritePrinter(printer, feed_and_cut_cmd)
    win32print.EndPagePrinter(printer)
    win32print.EndDocPrinter(printer)
    win32print.ClosePrinter(printer)
