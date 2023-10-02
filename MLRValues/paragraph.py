letterToWord = {1: 'a', 2: 'b', 3: 'c', 4: 'd', 5: 'e', 6: 'f', 7: 'g', 8: 'h', 9: 'i', 10: 'j', 11: 'k', 12: 'l', 13: 'm', 14: 'n', 15: 'o', 16: 'p', 17: 'q', 18: 'r', 19: 's', 20: 't', 21: 'u', 22: 'v', 23: 'w', 24: 'x', 25: 'y', 26: 'z'}
def adjustedRsquareLine(file, write_doc):
    
    modelSummary = file.tables[3]
    anova = file.tables[5]
    rSq = 2
    df = 2
    paragraph = ''
    while(rSq < len(modelSummary.rows) and df < len(anova.rows)):
        adjustedRsquare = str(round(float(modelSummary.rows[rSq].cells[3].text), 2))
        if adjustedRsquare[0] == '0':
            adjustedRsquare = adjustedRsquare[1:]
            if len(adjustedRsquare) != 3:
                    adjustedRsquare = adjustedRsquare + '0'

        elif adjustedRsquare[0:2] == '-0':
            adjustedRsquare = '-' + adjustedRsquare[2:]
            if len(adjustedRsquare) != 4:
                adjustedRsquare = adjustedRsquare + '0'


        rSq += 1
        dfReg = anova.rows[df].cells[3].text
        dfTot = anova.rows[df + 2].cells[3].text

        F = str(round(float(anova.rows[df].cells[5].text), 2))

        p = anova.rows[df].cells[6].text
        if not p[len(p) - 1] in '0123456789':
            p = p[:len(p) - 1]
        df += 3
        paragraph = paragraph + f"Model 1{letterToWord[rSq - 2]}; Adjusted R² = {adjustedRsquare}, F({dfReg}, {dfTot}) = {F}, p = {p}" + "\n"
        # write_doc.add_paragraph(f"Model 1{letterToWord[rSq - 2]}; Adjusted R² = {adjustedRsquare}, F({dfReg}, {dfTot}) = {F}, p = {p}")
    write_doc.add_paragraph(paragraph)
