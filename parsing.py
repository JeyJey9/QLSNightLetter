import pdfplumber
import re

def parse_pdf_values(pdf_path):
    """
    Extracts numeric segments from a single-page PDF summary containing "Actual Unit Concern".
    """
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
    target = next((ln for ln in text.split("\n") if "Actual Unit Concern" in ln), "")
    nums = [n for n in re.findall(r"\d+\.\d+|\d+", target) if n != '1000']
    if len(nums) < 14:
        return None
    mseg, wseg, dseg = nums[2:8], nums[8:14], nums[14:]
    pick = lambda seg, i: seg[-i] if len(seg) >= i else seg[-1]
    return {
        'monthly_last': pick(mseg,1), 'monthly_prev': pick(mseg,2),
        'weekly_last':  pick(wseg,1), 'weekly_prev':  pick(wseg,2),
        'daily_last':   pick(dseg,1), 'daily_prev':   pick(dseg,2),
        'daily_prev2':  pick(dseg,3),
    }

def extract_broken_pdf(pdf_path):
    """
    For multi-page PDFs: page 2 yields monthly, page 35 yields weekly/daily by fixed positions.
    """
    with pdfplumber.open(pdf_path) as pdf:
        # Monthly from page 2
        page2 = pdf.pages[1].extract_text() or ""
        line2 = next((ln for ln in page2.split("\n") if "Actual Unit Concern" in ln), "")
        nums2 = [n for n in re.findall(r"\d+\.\d+|\d+", line2) if n != '1000']
        if len(nums2) < 8:
            return None
        monthly_prev = nums2[6]
        monthly_last = nums2[7]

        # Weekly & daily from page 35
        text35 = pdf.pages[34].extract_text() or ""
        nums35 = []
        for ln in text35.split("\n"):
            tmp = re.findall(r"\d+\.\d+|\d+", ln)
            if len(tmp) >= 10:
                nums35 = tmp
                break
        if len(nums35) < 10:
            return None
        weekly_prev  = nums35[2]
        weekly_last  = nums35[3]
        daily_prev2  = nums35[7]
        daily_prev   = nums35[8]
        daily_last   = nums35[9]

    return {
        'monthly_last': float(monthly_last),
        'monthly_prev': float(monthly_prev),
        'weekly_last':  float(weekly_last),
        'weekly_prev':  float(weekly_prev),
        'daily_last':   float(daily_last),
        'daily_prev':   float(daily_prev),
        'daily_prev2':  float(daily_prev2),
    }
