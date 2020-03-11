import sys
from io import StringIO
import tokenize
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE, FORMULA


def getTokens(formula):
    result = list(tokenize.generate_tokens(StringIO(formula).readline))
    if result[0].string == "=":
        result.pop(0)
    return fixNames(result)

def fixNames(tokenList):
    """I'm co-opting the python tokenizer and have to tweak the NAME tokens
    to allow for spreadsheet references.
    A custom tokenizer or, even better, one for calc sheets would obviate
    this completely.
    """
    result = []
    while len(tokenList) > 0:
        token = tokenList.pop(0)
        if token.type in (tokenize.ERRORTOKEN, tokenize.COMMENT) \
            or (token.type == tokenize.OP and token.string in ('.',':')):
           # if previous isn't an OP, have to bind to that
           # ditto when following.
           if len(result) > 0:
               prv = result.pop()
               if prv.type == tokenize.OP:
                   result.append(prv) # put it back
               else:
                   token = tokenize.TokenInfo(tokenize.NAME, prv.string+token.string, prv.start, token.end, token.line)

           nxt = tokenList.pop(0) if len(tokenList)>0 else None
           while nxt != None and (nxt.type not in (tokenize.OP, tokenize.NEWLINE) or nxt.string == '.'):
               token = tokenize.TokenInfo(tokenize.NAME, token.string+nxt.string, token.start, nxt.end, token.line)
               nxt = tokenList.pop(0) if len(tokenList)>0 else None

           if nxt != None: tokenList.insert(0, nxt)

        result.append(token)

    return result
           
def substitueFormulaReference(token):
    """If a NAME token corresponds to a formula rather than a value,
    replace the NAME with the respective formula taking care the context
    of resulting NAMEs is preserved e.g.
    $Sheet1.A1 = A3 * A2 --> $Sheet1.A3 * $Sheet1.A2
    """
    if token == None or token.type != tokenize.NAME: return [token]
    formula = hasFormula(token.string)
    if formula == None: return [token]
    result = []
    formula = tokenize.untokenize(getTokens(formula)) # expunge leading =
    if '.' in token.string: # qualified name
        sheet, cell = token.string.split('.')
        for tok in getTokens('(' + formula + ')'):
            if tok.type == tokenize.NAME and '.' not in tok.string:
                result.append(tokenize.TokenInfo(tok.type, sheet + '.' + tok.string, tok.start, tok.end, tok.line))
            else:
                result.append(tok)
    else:
        result.extend(getTokens('(' + formula + ')'))
    return result[:-1] # drop end marker

def processPrecedents(tokens):
    """Process the tokenized formula applying precendent substitution"""
    i = 0
    result = tokens.copy()
    while i < len(result):
        if result[i].type == tokenize.NAME:
            name = result[i]
            result.remove(name)
            result[i:i] = substitueFormulaReference(name)
            if name.string == result[i].string: i = i + 1
        else:
            i = i + 1
    return result


def getFormula(tokens):
    """Strip tokens back as a string"""
    return " ".join(map(lambda x:x.string, tokens))

def hasFormula(name):
    """Check if given name corresponds to a cell containing a formula and if so, return the formula
    Complicated by the fact that the name could refer to a cell
    in a different file. Not going to follow those, so treat as if atomic value.
    Could be a range or anything the tokenizer sees as a NAME e.g. functions
    """
    msgBox("hasFormula", name)
    # Tokenize the name string - if 2nd token is a comment --> in different file
    tokens = getTokens(name)
    if len(tokens) > 2 and tokens[1].type == tokenize.COMMENT:
        return None
    refParts = name.split('.')
    if ':' in refParts[-1]: return None #range, baby, range
    if len(refParts) == 1: # simple cell reference, no dots
        msgBox("hasFormula simple ref", name)
        try:
            cell = model.CurrentController.ActiveSheet.getCellRangeByName(name)
        except:
            return None
    else:
        sheetName = '.'.join(refParts[:-1]).replace('$','').strip()
        cellName = refParts[-1]
        msgBox("hasFormula resolve sheet", sheetName + "\n" + cellName)
        try:
            sheet = model.Sheets.getByName(sheetName)
            cell = sheet.getCellRangeByName(cellName)
        except:
            return None

    if cell.getType() == FORMULA: return cell.Formula
    return None


def msgBox(title, content):
    if debugMeBaby == 0: return None
    parentwin = model.CurrentController.Frame.ContainerWindow
    box = parentwin.getToolkit().createMessageBox(parentwin, MESSAGEBOX,  BUTTONS_OK, title, content)
    result = box.execute()
    if result == OK:
        print("OK")
    return None

def FlattenFormula(*args):
    """Flatten precedents in formula in a single expression.
    Examines active cell for a formula and shows the transformed
    formula. Cell references are replaced by the formula
    they contain recursively. Cell references that refer to a
    value, range or other file are left alone."""
    desktop = XSCRIPTCONTEXT.getDesktop()
    global model 
    model = desktop.getCurrentComponent()
    if not hasattr(model, "Sheets"):
        return None
    active_sheet = model.CurrentController.ActiveSheet
    active_cell = model.CurrentController.getSelection()
    if active_cell.getType() != FORMULA:
        return None
    flat = getFormula(processPrecedents(getTokens(active_cell.Formula)))
    resultStr = "Current cell: {},{}\nOriginal Formula : {}\nFlattened:\n{}".format(
        active_cell.getCellAddress().Column,
        active_cell.getCellAddress().Row,
        active_cell.Formula,
        flat)

    parentwin = model.CurrentController.Frame.ContainerWindow
    box = parentwin.getToolkit().createMessageBox(parentwin, MESSAGEBOX,  BUTTONS_OK, "Flatten formula result", resultStr)
    result = box.execute()
    if result == OK:
        print("OK")


debugMeBaby = 0
g_exportedScripts = FlattenFormula, 
