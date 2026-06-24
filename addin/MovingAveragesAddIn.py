# =====================================================================
# MovingAveragesAddIn.py
#
# A LibreOffice Calc Add-In (UNO component) implementing the MOVAVG
# spreadsheet function in pure Python. Unlike a Basic macro, an Add-In
# function autocompletes as you type and appears in the Function Wizard
# (its name / argument descriptions come from CalcAddIns.xcu).
#
# Function:
#   MOVAVG(data, period, [type])
#       type "s" / omitted -> simple moving average
#       type "e"           -> exponential moving average
#   Returns the moving average value at the END of <data>. Anchor the
#   start of the range and fill down, e.g. =MOVAVG(A$2:A2, 5).
#
# Dependencies: none beyond LibreOffice's bundled Python + UNO bridge.
# =====================================================================

import unohelper

from com.sun.star.lang import XServiceInfo, XServiceName
from org.davidjackson.smaema import XMovingAverages

# The implementation name is private to this component; the two service
# names are what LibreOffice and the CalcAddIns.xcu look the add-in up by.
IMPL_NAME = "org.davidjackson.smaema.MovingAveragesAddIn.python"
OWN_SERVICE = "org.davidjackson.smaema.MovingAveragesAddIn"
ADDIN_SERVICE = "com.sun.star.sheet.AddIn"


def _to_double(v):
    """Safe numeric conversion; blanks / text become 0.0."""
    try:
        if v is None or v == "":
            return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _flatten(data):
    """Normalize a scalar / 1-D / 2-D (cell range) argument into a flat
    list of floats. Cell ranges arrive as a sequence of rows."""
    out = []
    if isinstance(data, (list, tuple)):
        for row in data:
            if isinstance(row, (list, tuple)):
                for cell in row:
                    out.append(_to_double(cell))
            else:
                out.append(_to_double(row))
    else:
        out.append(_to_double(data))
    return out


def _movavg(data, period, kind):
    period = int(period)
    if period < 1:
        return "Error: period must be >= 1"

    k = "s"
    if kind is not None:
        s = str(kind).strip().lower()
        if s:
            k = s

    values = _flatten(data)
    n = len(values)
    if n < period:
        return ""  # not enough data yet

    if k[:1] == "e":
        # Exponential: seed with the simple average of the first <period>
        # values, then roll forward to the end of the range.
        prev = sum(values[:period]) / period
        alpha = 2.0 / (period + 1)
        for i in range(period, n):
            prev = alpha * values[i] + (1.0 - alpha) * prev
        return prev

    # Simple: trailing average of the last <period> values.
    return sum(values[n - period:n]) / period


class MovingAveragesAddIn(unohelper.Base, XMovingAverages, XServiceInfo, XServiceName):
    def __init__(self, ctx):
        self.ctx = ctx

    # --- XMovingAverages (the spreadsheet functions) ---
    def MOVAVG(self, data, period, kind):
        return _movavg(data, period, kind)

    # --- XServiceName ---
    def getServiceName(self):
        return OWN_SERVICE

    # --- XServiceInfo ---
    def getImplementationName(self):
        return IMPL_NAME

    def supportsService(self, name):
        return name in (ADDIN_SERVICE, OWN_SERVICE)

    def getSupportedServiceNames(self):
        return (ADDIN_SERVICE, OWN_SERVICE)


# Passive registration: LibreOffice imports this module and reads
# g_ImplementationHelper to learn what this component provides.
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MovingAveragesAddIn, IMPL_NAME, (ADDIN_SERVICE, OWN_SERVICE),
)
