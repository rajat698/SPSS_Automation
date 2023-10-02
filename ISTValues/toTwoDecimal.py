#This function converts a string containing a floating point integer to a string
#with same float rounded up to two decimle places

def toTwoDecimal(string):
    result = str(round(float(string), 2))
    return result