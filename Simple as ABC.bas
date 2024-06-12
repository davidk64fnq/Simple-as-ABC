    global gOpenfile$, gFontString$, gNumbers$, gVector$, gHighlightRowCol,_
           gMainUpperLeftX, gMainUpperLeftY, gMainWindowHeight, gMainWindowWidth,_          
           gCheckerWndOpen, gCluesWndOpen, gStatesWndOpen, gGameChanged, gTrue, gFalse

    ' There is a grid of 6 by 6 cells. Each cell consists of 6 buttons, a row of 5 small
    ' buttons across the top and 1 large button below. The middle 3 small buttons are used
    ' for recording which of A, B, and C the user thinks the cell MIGHT be. The big button
    ' records which of A, B, and C the user decides the cell ACTUALLY is. User cycles through
    ' A, B, C and blank with successive left mouse clicks of large button. Clicking on small
    ' buttons toggles between a single letter and blank.

    call initGlobals
    call initArrays
    call loadBitmapFiles
    call createButtons  ' Separate sub so can be put at bottom of this file out of the way
    call openMainWnd

    sub initGlobals
        gOpenfile$ = ""         ' Currently open file for save menu option
        gFontString$ = "font ms_sans_serif 10"
        gNumbers$ = "zero one two three four five six"
        gVector$ = ""           ' Either a row or a column extracted from grid as string of 6 letters
        gHighlightRowCol = -1   ' Number of row/col currently highlighted in main window
        gCheckerWndOpen = gFalse
        gCluesWndOpen = gFalse
        gStatesWndOpen = gFalse
        gGameChanged = gFalse   ' To check if user wants to save on program close
        gTrue = 1
        gFalse = 0
    end sub

    ' This program use array indexes from 1 so dimensions are 1 larger than required number of elements.
    ' Top line of cells is row 1, lefthand column of cells is column 1.
    sub initArrays
        ' Related to "Main" program window - the grid of cells
        dim largeLabels$(7,7)   ' A, B, C or ?(displayed as no letter)
        dim smallLabels$(7,42)  ' A/?, B/?, and C/? (3 dimensional array 6x6x5)
        for row = 1 to 6
            for col = 1 to 6
                largeLabels$(row, col) = "?"
                for small = 1 to 5
                    select case small
                        case 1, 5
                            smallLabels$(row, col * 6 + small) = "?"
                        case 2
                            smallLabels$(row, col * 6 + small) = "A"
                        case 3
                            smallLabels$(row, col * 6 + small) = "B"
                        case 4
                            smallLabels$(row, col * 6 + small) = "C"
                    end select
                next small
            next col
        next row
        dim info$(10,10)       ' Used to test file existance

        ' Related to "Checker" window
        dim puzzleInfo$(13,4)  ' Store new puzzle in format useful for programmatic tasks
                               ' 1st dimension clue no. 1 to 12
                               ' 2nd dimension: index 1 is clue template, index 2 is X value, index 3 is Y value

        ' Related to "States" window
        dim stateChecked$(11)               ' Which state in state window currently checked
        dim stateDescriptions$(11)
        dim stateLargeLabels$(11,49)        ' 3 dim array 10x6x6 stores A/B/C/?
        dim stateSmallLabels$(11,49)        ' 3 dim array 10x6x6 stores string showing small buttons
                                            ' eg "A?C" is A and C set, B off
        dim stateClueConditionMet$(11,13)
        for state = 1 to 10
            stateChecked$(state) = "0"
            stateDescriptions$(state) = "empty"
            for row = 1 to 6
                for col = 1 to 6
                    stateLargeLabels$(state, row * 7 + col) = "?"
                    stateSmallLabels$(state, row * 7 + col) = "???"
                next col
            next row
            for clueNo = 1 to 12
                stateClueConditionMet$(state, clueNo) = "0"
            next clueNo
        next state

        ' Related to "Clues" windows
        dim currentClues$(13)       ' Used to store clues for saving to file purposes
                                    ' index 1 to 6 are row clues and 7 to 12 column clues
                                    ' note some puzzles have less than 12 clues
        dim clueConditionMet$(13)   ' Which clues ticked by user as having been satisfied
        for clueNo = 1 to 12
            currentClues$(clueNo) = ""
            clueConditionMet$(clueNo) = "0"
        next clueNo

        dim clueTemplates$(7)
        clueTemplates$(1) = "The X's are in adjacent squares"
        clueTemplates$(2) = "The X's are beyond the Y's"
        clueTemplates$(3) = "The X's are between the Y's"
        clueTemplates$(4) = "Each X is next to and beyond a Y"
        clueTemplates$(5) = "No two adjacent squares contain the same letter"
        clueTemplates$(6) = "Any 3 consecutive squares contain 3 different letters"
    end sub

    ' All needed bmpbutton bitmaps are packaged with the program as external files and loaded into
    ' memory when the program starts. There are two versions of each bitmap one being for when a row
    ' or column is highlighted.
    '
    ' Naming convention for bitmaps is
    '   First character: 1 to 6 showing button position (1 to 5 are small left to right, 6 is big below)
    '   Second character: N = unhighlighted H = highlighted
    '   Third character: E = no letter A, B, C = the respective letter
    sub loadBitmapFiles
        states$ = "N H"
        letters$ = "A B C E"

        for stateNo = 1 to 2
            stateChar$ = word$(states$, stateNo)
            bmpName$ = "1" + stateChar$ + "E"
            call loadBitmap bmpName$
            for buttonNo = 2 to 4
                for letter = 1 to 4
                    bmpName$ = str$(buttonNo); stateChar$; word$(letters$, letter)
                    call loadBitmap bmpName$
                next letter
            next buttonNo
            bmpName$ = "5" + stateChar$ + "E"
            call loadBitmap bmpName$
            for letter = 1 to 4
                bmpName$ = "6" + stateChar$ + word$(letters$, letter)
                call loadBitmap bmpName$
            next letter
        next stateNo
    end sub

    sub loadBitmap bmpName$
        filename$ = "bmp\" + bmpName$ + ".bmp"
        if fileExists(DefaultDir$, filename$) then
            loadbmp bmpName$, filename$
        else
            notice "Fatal Error!" + chr$(13)_
                    + chr$(34) + filename$ + chr$(34) + " is missing." + chr$(13) + chr$(13)_
                    + "You may need to run " + chr$(34) + "Create Bitmap files.bas" + chr$(34) + chr$(13)_
                    + "to generate the needed bitmap files"
            end
        end if
    end sub

    sub openMainWnd
        nomainwin
        menu #mainWnd, "&File", "&New", handleNew, "&Open...", handleOpen, "&Save", handleSave,_
                       "Save &As...", handleSaveAs
        menu #mainWnd, "&View", "&Clues", openCluesWnd, "S&olution Checker", openCheckerWnd, "&States", openStatesWnd
        call setMainWndSize
        open "Simple as A, B, C?" for window_nf as #mainWnd
        #mainWnd "trapclose closeMainWnd"
        call refreshBitmaps
        gGameChanged = gFalse
        call openCluesWnd
        call openStatesWnd
        call openCheckerWnd
        wait
    end sub

    sub closeMainWnd handle$
        ' If current game has changed give user chance to save first
        if gGameChanged then
            confirm "Save changes?"; save$
            if save$ = "yes" then call handleSave
        end if

        confirm "Do you want to quit " + chr$(34) + "Simple as A, B, C?" + chr$(34); quit$
        if quit$ = "no" then wait
        call unloadBitmaps
        if gCluesWndOpen then close #cluesWnd
        if gCheckerWndOpen then close #checkerWnd
        if gStatesWndOpen then close #statesWnd
        close #handle$
        end
    end sub

    sub unloadBitmaps
        states$ = "N H"
        letters$ = "A B C E"

        for stateNo = 1 to 2
            stateChar$ = word$(states$, stateNo)
            bmpName$ = "1" + stateChar$ + "E"
            unloadBmp bmpName$
            for buttonNo = 2 to 4
                for letter = 1 to 4
                    bmpName$ = str$(buttonNo); stateChar$; word$(letters$, letter)
                    unloadBmp bmpName$
                next letter
            next buttonNo
            bmpName$ = "5" + stateChar$ + "E"
            unloadBmp bmpName$
            for letter = 1 to 4
                bmpName$ = "6" + stateChar$ + word$(letters$, letter)
                unloadBmp bmpName$
            next letter
        next stateNo
    end sub

    ' Format for puzzle files is 1 puzzle per line, only rows or columns having a clue are
    ' listed, clues comma separated with no whitespace.
    '   e.g. 1.3ac is row 1, template clue 3 (see initArrays), X = "A", Y = "C"
    '        9.4ca is column 3 (9 - 6 = 3), template clue 4, X = "C", Y = "A"
    sub handleNew
        ' If current game has changed give user chance to save first
        if gGameChanged then
            confirm "Save changes?"; save$
            if save$ = "yes" then
                call handleSave
            end if
        end if

        ' Open file
        filedialog "Open puzzle file", "*.puz", fileName$
        if fileName$ = "" then exit sub
        gOpenfile$ = ""
        open fileName$ for input as #inFile

        ' Prompt for which puzzle in file to load
        dim puzzles$(1000)
        while eof(#inFile) = gFalse
            lineCount = lineCount + 1
            line input #inFile, puzzles$(lineCount)
        wend
        close #inFile
        prompt "Which puzzle no. do you wish to load?" + chr$(13)_
               + "An integer between 1 and " + str$(lineCount); response$
        puzzleNo = int(val(response$))
        if (puzzleNo < 1) or (puzzleNo > lineCount) then
            notice "Puzzle number " + str$(puzzleNo) + " is not in the range 1 to " + str$(lineCount)
            exit sub
        end if

        ' Parse selected puzzle string
        call initArrays
        clueInvalid = gFalse
        for index = 1 to 12
            xValue$ = ""
            yValue$ = ""
            nextClue$ = word$(puzzles$(puzzleNo), index, ",")
            if nextClue$ = "" then exit for

            ' Get clue number (1 to 12) which is integer on lefthand side of dot
            lhsDot$ = word$(nextClue$, 1, ".")
            clueNo = val(lhsDot$)
            if (clueNo <> int(clueNo)) or (clueNo < 1) or (clueNo > 12) then clueInvalid = gTrue

            ' Get clue template (1 to 6) which is integer on righthand side of dot
            rhsDot$ = word$(nextClue$, 2, ".")
            clueTemplateNo = val(left$(rhsDot$, 1))
            puzzleInfo$(clueNo, 1) = str$(clueTemplateNo)
            if (clueTemplateNo <> int(clueTemplateNo)) or (clueTemplateNo < 1) or (clueTemplateNo > 6) then_
                clueInvalid = gTrue

            ' Get "X" character if present, 1st char after clue template number (one of [ABCabc])
            if len(rhsDot$) > 1 then
                xValue$ = upper$(mid$(rhsDot$, 2,1))
                puzzleInfo$(clueNo, 2) = xValue$
                if not(instr("ABC", xValue$)) then clueInvalid = gTrue
            end if

            ' Get "Y" character if present, 2nd char after clue template number (one of [ABCabc])
            if len(rhsDot$) > 2 then
                yValue$ = upper$(mid$(rhsDot$, 3,1))
                puzzleInfo$(clueNo, 3) = yValue$
                if not(instr("ABC", yValue$)) then clueInvalid = gTrue
            end if

            ' Construct clue
            clue$ = clueTemplates$(clueTemplateNo)
            if instr(clue$, "X") then
                if (xValue$ <> "") then
                    clue$ = strReplace$(clue$, "X", xValue$)
                else
                    clueInvalid = gTrue
                end if
            end if
            if instr(clue$, "Y") then
                if (yValue$ <> "") then
                    clue$ = strReplace$(clue$, "Y", yValue$)
                else
                    clueInvalid = gTrue
                end if
            end if
            if clueNo <= 6 then
                clue$ = strReplace$(clue$, "beyond", "further right than")
            else
                clue$ = strReplace$(clue$, "beyond", "below")
            end if
            clue$ = strReplace$(clue$, "a A", "an A")
            currentClues$(clueNo) = clue$
        next index
        if clueInvalid then
            notice "A problem was encountered reading this puzzle from file." + chr$(13) +_
                   "Refer to the help for more info on puzzle file format."
            exit sub
        end if

        call refreshBitmaps
        if gCluesWndOpen then call refreshClues
        stateChecked$(1) = "1"
        if gStatesWndOpen then call refreshStates
        gGameChanged = gFalse
    end sub

    sub handleOpen
        ' If current game has changed give user chance to save first
        if gGameChanged then
            confirm "Save changes?"; save$
            if save$ = "yes" then
                call handleSave
                exit sub
            end if
        end if

        ' Open file
        filedialog "Open previously saved game progress", "*.prg", fileName$
        if fileName$ = "" then exit sub
        gOpenfile$ = fileName$
        open fileName$ for input as #inFile

        ' Load "View Clues" window checkbox set/unset
        for clueNo = 1 to 12
            clueConditionMet$(clueNo) = inputto$(#inFile, ",")
        next clueNo
        line input #inFile, dummy$  ' Advancing to next line past trailing ","

        ' Load clues
        for clueNo = 1 to 12
            line input #inFile, currentClues$(clueNo)
        next clueNo

        ' Load large and small labels
        for row = 1 to 6
            for col = 1 to 6
                largeLabels$(row, col) = inputto$(#inFile, ",")
                for small = 1 to 5
                    smallLabels$(row, col * 6 + small) = inputto$(#inFile, ",")
                next small
            next col
        next row
        line input #inFile, dummy$ 

        ' Load state checkboxes
        for state = 1 to 10
            stateChecked$(state) = inputto$(#inFile, ",")
        next state
        line input #inFile, dummy$ 

        ' Load state descriptions
        for state = 1 to 10
            line input #inFile, stateDescriptions$(state)
        next state

        ' Load states
        for state = 1 to 10
            for row = 1 to 6
                for col = 1 to 6
                    stateLargeLabels$(state, row * 7 + col) = inputto$(#inFile, ",")
                    stateSmallLabels$(state, row * 7 + col) = inputto$(#inFile, ",")
                next col
            next row
            line input #inFile, dummy$ 
        next state

        ' Load state clue conditions met checkboxes
        for state = 1 to 10
            for clueNo = 1 to 12
                stateClueConditionMet$(state, clueNo) = inputto$(#inFile, ",")
            next clueNo
            line input #inFile, dummy$ 
        next state

        ' Load puzzleInfo$ - the coded clues used in checker window subs
        for clueNo = 1 to 12
            for itemNo = 1 to 3
                puzzleInfo$(clueNo, itemNo) = inputto$(#inFile, ",")
            next itemNo
            line input #inFile, dummy$ 
        next clueNo

        close #inFile
        call refreshBitmaps
        if gCluesWndOpen then call refreshClues
        if gStatesWndOpen then call refreshStates
        gGameChanged = gFalse
    end sub

    sub handleSave
        if gOpenfile$ = "" then
            call handleSaveAs
            exit sub
        end if

        open gOpenfile$ for output as #outFile
        call saveProgress "#outFile"
        close #outFile
    end sub

    sub handleSaveAs
        ' Get file name to save to
        filedialog "Save game progress", "*.prg", fileName$
        if fileName$ = "" then
            exit sub
        end if

        ' Confirm overwrite if it exists already
        if fileExists(DefaultDir$, fileName$) then
            c$ = "This will overwrite existing file contents!" + chr$(13) + chr$(13)
            c$ = c$ + "Continue?"
            confirm c$; answer$
            if answer$ = "no" then exit sub
        end if

        gOpenfile$ = fileName$
        open fileName$ for output as #outFile
        call saveProgress "#outFile"
        close #outFile
    end sub

    sub saveProgress outFile$
        ' Save "View Clues" window checkbox set/unset
        for clueNo = 1 to 12
            print #outFile$, clueConditionMet$(clueNo) + ",";
        next clueNo
        print #outFile$, ""

        ' Save clues
        for clueNo = 1 to 12
            print #outFile$, currentClues$(clueNo)
        next clueNo

        ' Save large and small labels
        for row = 1 to 6
            for col = 1 to 6
                print #outFile$, largeLabels$(row, col) + ",";
                for small = 1 to 5
                    print #outFile$, smallLabels$(row, col * 6 + small) + ",";
                next small
            next col
        next row
        print #outFile$, ""

        ' Save state checkboxes
        for state = 1 to 10
            print #outFile$, stateChecked$(state) + ",";
        next state
        print #outFile$, ""

        ' Save states descriptions
        for state = 1 to 10
            print #outFile$, stateDescriptions$(state)
        next state

        ' Save states
        for state = 1 to 10
            for row = 1 to 6
                for col = 1 to 6
                    print #outFile$, stateLargeLabels$(state, row * 7 + col) + ",";
                    print #outFile$, stateSmallLabels$(state, row * 7 + col) + ",";
                next col
            next row
            print #outFile$, ""
        next state

        ' Save state clue conditions met checkboxes
        for state = 1 to 10
            for clueNo = 1 to 12
                print #outFile$, stateClueConditionMet$(state, clueNo) + ",";
            next clueNo
            print #outFile$, ""
        next state

        ' Save puzzleInfo$ - the coded clues used in checker window subs
        for clueNo = 1 to 12
            for itemNo = 1 to 3
                print #outFile$, puzzleInfo$(clueNo, itemNo) + ",";
            next itemNo
            print #outFile$, ""
        next clueNo

        gGameChanged = gFalse
    end sub

    ' Cycling through A, B, C and no letter when user clicks large button
    sub handleLargeClick buttonhandle$
        clickedRow = val(mid$(buttonhandle$, 14, 1))
        clickedCol = val(mid$(buttonhandle$, 15, 1))
        ' Prevent user from editing rows and columns other than selected row/col
        if ((gHighlightRowCol >= 1) and  (gHighlightRowCol <= 6) and (gHighlightRowCol <> clickedRow)) or _
           ((gHighlightRowCol > 6) and (gHighlightRowCol <> 6 + clickedCol)) then exit sub
        select case largeLabels$(clickedRow, clickedCol)
            case "?"
                largeLabels$(clickedRow, clickedCol) = "A"
            case "A"
                largeLabels$(clickedRow, clickedCol) = "B"
            case "B"
                largeLabels$(clickedRow, clickedCol) = "C"
            case "C"
                largeLabels$(clickedRow, clickedCol) = "?"
        end select
        gGameChanged = gTrue
        call refreshBitmaps
    end sub

    ' Toggle between (A, B or C) and no letter when user clicks one of middle three small buttons
    sub handleSmallClick buttonhandle$
        clickedRow = val(mid$(buttonhandle$, 14, 1))
        clickedCol = val(mid$(buttonhandle$, 15, 1))
        clickedSmall = val(mid$(buttonhandle$, 17, 1))
        ' Prevent user from editing rows and columns other than selected row/col
        if ((gHighlightRowCol <= 6) and (gHighlightRowCol <> clickedRow)) or _
           ((gHighlightRowCol > 6) and (gHighlightRowCol <> 6 + clickedCol)) then exit sub
        select case smallLabels$(clickedRow, clickedCol * 6 + clickedSmall)
            case "?"
                select case clickedSmall
                    case 2
                        smallLabels$(clickedRow, clickedCol * 6 + clickedSmall) = "A"
                    case 3
                        smallLabels$(clickedRow, clickedCol * 6 + clickedSmall) = "B"
                    case 4
                        smallLabels$(clickedRow, clickedCol * 6 + clickedSmall) = "C"
                end select
            case "A", "B", "C"
                smallLabels$(clickedRow, clickedCol * 6 + clickedSmall) = "?"
        end select
        gGameChanged = gTrue
        call refreshBitmaps
    end sub

    sub refreshBitmaps
        for row = 1 to 6
            for col = 1 to 6
                ' Refresh large button bitmaps
                letter$ = largeLabels$(row, col)
                if largeLabels$(row, col) = "?" then
                    letter$ = "E"
                end if
                if ((row = gHighlightRowCol) or (col = gHighlightRowCol - 6)) then
                    bmp$ = "6H" + letter$
                else
                    bmp$ = "6N" + letter$
                end if
                buttonhandle$ = "#mainWnd.cell" + str$(row) + str$(col)
                #buttonhandle$ "bitmap "; bmp$

                ' Refresh small button bitmaps
                for small = 1 to 5
                    letter$ = smallLabels$(row, col * 6 + small)
                    if smallLabels$(row, col * 6 + small) = "?" then
                        letter$ = "E"
                    end if
                    if ((row = gHighlightRowCol) or (col = gHighlightRowCol - 6)) then
                        bmp$ = str$(small) + "H" + letter$
                    else
                        bmp$ = str$(small) + "N" + letter$
                    end if
                    buttonhandle$ = "#mainWnd.cell" + str$(row) + str$(col) + "." + str$(small)
                    #buttonhandle$ "bitmap "; bmp$
                next small
            next col
        next row
    end sub

    sub setMainWndSize
        WindowWidth = 200
        WindowHeight = 200
        menu #gr, "Dummy"
        open "Getting window size" for graphics_nsb_nf as #gr
        #gr, "home ; down ; posxy penX penY"
        WindowWidth = 600 + 200 - (2 * penX)
        gMainWindowWidth = WindowWidth
        WindowHeight = 600 + 200 - (2 * penY)
        gMainWindowHeight = WindowHeight
        UpperLeftX = DisplayWidth/2 - WindowWidth/2
        gMainUpperLeftX = UpperLeftX
        UpperLeftY = DisplayHeight/2 - WindowHeight/2
        gMainUpperLeftY = UpperLeftY
        close #gr
    end sub

    '--------------------------------------------------------------------------------------------------------------
    '
    '                   "Clues" Window
    '
    '--------------------------------------------------------------------------------------------------------------

    ' Shows the set of clues for the current puzzle. A column of check boxes on the left are
    ' there so user can tick a clue to remind them the clue constraints have been satisfied.
    ' The clues are buttons. If a clue button is pressed it toggles highlighting (bold text) for the
    ' corresponding row or column.
    sub openCluesWnd
        WindowWidth = 475
        WindowHeight = gMainWindowHeight
        UpperLeftX = gMainUpperLeftX - 475
        UpperLeftY = gMainUpperLeftY
        groupbox #cluesWnd.groupbox1, "Across", 26, 16, 415, 285
        checkbox #cluesWnd.checkboxR1, "1", handleCluesWndCBset, handleCluesWndCBunset, 51, 51, 24, 20
        checkbox #cluesWnd.checkboxR2, "2", handleCluesWndCBset, handleCluesWndCBunset, 51, 91, 24, 20
        checkbox #cluesWnd.checkboxR3, "3", handleCluesWndCBset, handleCluesWndCBunset, 51, 131, 24, 20
        checkbox #cluesWnd.checkboxR4, "4", handleCluesWndCBset, handleCluesWndCBunset, 51, 171, 24, 20
        checkbox #cluesWnd.checkboxR5, "5", handleCluesWndCBset, handleCluesWndCBunset, 51, 211, 24, 20
        checkbox #cluesWnd.checkboxR6, "6", handleCluesWndCBset, handleCluesWndCBunset, 51, 251, 24, 20
        button #cluesWnd.buttonR1, "", handleCluesWndButClick, UL, 91, 46, 340, 25
        button #cluesWnd.buttonR2, "", handleCluesWndButClick, UL, 91, 86, 340, 25
        button #cluesWnd.buttonR3, "", handleCluesWndButClick, UL, 91, 126, 340, 25
        button #cluesWnd.buttonR4, "", handleCluesWndButClick, UL, 91, 166, 340, 25
        button #cluesWnd.buttonR5, "", handleCluesWndButClick, UL, 91, 206, 340, 25
        button #cluesWnd.buttonR6, "", handleCluesWndButClick, UL, 91, 246, 340, 25
        groupbox #cluesWnd.groupbox2, "Down", 26, 311, 415, 285
        checkbox #cluesWnd.checkboxC1, "1", handleCluesWndCBset, handleCluesWndCBunset, 51, 346, 24, 20
        checkbox #cluesWnd.checkboxC2, "2", handleCluesWndCBset, handleCluesWndCBunset, 51, 386, 24, 20
        checkbox #cluesWnd.checkboxC3, "3", handleCluesWndCBset, handleCluesWndCBunset, 51, 426, 24, 20
        checkbox #cluesWnd.checkboxC4, "4", handleCluesWndCBset, handleCluesWndCBunset, 51, 466, 24, 20
        checkbox #cluesWnd.checkboxC5, "5", handleCluesWndCBset, handleCluesWndCBunset, 51, 506, 24, 20
        checkbox #cluesWnd.checkboxC6, "6", handleCluesWndCBset, handleCluesWndCBunset, 51, 546, 24, 20
        button #cluesWnd.buttonC1, "", handleCluesWndButClick, UL, 91, 341, 340, 25
        button #cluesWnd.buttonC2, "", handleCluesWndButClick, UL, 91, 381, 340, 25
        button #cluesWnd.buttonC3, "", handleCluesWndButClick, UL, 91, 421, 340, 25
        button #cluesWnd.buttonC4, "", handleCluesWndButClick, UL, 91, 461, 340, 25
        button #cluesWnd.buttonC5, "", handleCluesWndButClick, UL, 91, 501, 340, 25
        button #cluesWnd.buttonC6, "", handleCluesWndButClick, UL, 91, 541, 340, 25
        open "View Clues" for window_nf as #cluesWnd
        gCluesWndOpen = gTrue
        #cluesWnd gFontString$
        #cluesWnd "trapclose closeCluesWnd"
        call refreshClues
        scan
    end sub

    sub closeCluesWnd handle$
        close #handle$
        gCluesWndOpen = gFalse
    end sub

    sub handleCluesWndCBset handle$
        ' What clue number has been clicked
        clueNo = val(right$(handle$, 1))
        rowOrCol$ = mid$(handle$, 19, 1)
        if rowOrCol$ = "C" then clueNo = clueNo + 6

        clueConditionMet$(clueNo) = "1"
    end sub

    sub handleCluesWndCBunset handle$
        ' What clue number has been clicked
        clueNo = val(right$(handle$, 1))
        rowOrCol$ = mid$(handle$, 19, 1)
        if rowOrCol$ = "C" then clueNo = clueNo + 6

        clueConditionMet$(clueNo) = "0"
    end sub

    ' Clicking the clue buttons in view clue window toggle highlighting of the corresponding
    ' row or column. gHighlightRowCol records the currently selected row/col
    ' if there is one. "-1" means there is no corresponding row/col highlighted
    sub handleCluesWndButClick handle$
        ' Has user clicked row or column button and what clue number?
        rowOrCol$ = mid$(handle$, 17, 1)
        rowOrColNo = val(mid$(handle$, 18, 1))

        if rowOrCol$ = "C" then rowOrColNo = rowOrColNo + 6

        if gHighlightRowCol = rowOrColNo then
            ' row/col already highlighted so toggle off
            gHighlightRowCol = -1
        else
            ' another row/col highlighted or none are so toggle clicked row/col highlighting on
            gHighlightRowCol = rowOrColNo
        end if
        call refreshBitmaps
        call refreshClues
    end sub

    sub refreshClues
        if gCluesWndOpen = gFalse then exit sub
        rowOrCol$ = "R C"
        for rowOrCol = 1 to 2
            for clueNo = 1 to 6
                ' Refresh which clue if any is highlighted
                buttonhandle$ = "#cluesWnd.button" + word$(rowOrCol$, rowOrCol) + str$(clueNo)
                clueNo1to12 = (rowOrCol - 1) * 6 + clueNo
                #buttonhandle$ currentClues$(clueNo1to12)
                if (clueNo1to12 = gHighlightRowCol) then
                   #buttonhandle$ "!font " + gFontString$ + " bold"
                else
                   #buttonhandle$ "!font " + gFontString$
                end if

                ' Refresh which clues have checkbox ticked
                CBhandle$ = "#cluesWnd.checkbox" + word$(rowOrCol$, rowOrCol) + str$(clueNo)
                if clueConditionMet$((rowOrCol - 1) * 6 + clueNo) = "1" then
                    #CBhandle$ "set"
                else
                    #CBhandle$ "reset"
                end if
            next clueNo
        next rowOrCol
    end sub

    '--------------------------------------------------------------------------------------------------------------
    '
    '                   "Solution Checker" Window
    '
    '--------------------------------------------------------------------------------------------------------------

    sub openCheckerWnd
        WindowWidth = 410
        WindowHeight = gMainWindowHeight - 465  ' 465 is state window height
        UpperLeftX = gMainUpperLeftX + gMainWindowWidth
        UpperLeftY = DisplayHeight/2 - gMainWindowHeight/2 + 465
        texteditor #checkerWnd.messagesTE, 20, 21, 360, 80
        button #checkerWnd.checkSolnB, "Check Solution", handleSolvedQuery, UL, 281, 115, 100, 25
        statictext #checkerWnd.stepDelayST, "Message Delay (sec)", 26, 119, 130, 20
        textbox #checkerWnd.stepDelayTB, 166, 117, 25, 25
        open "Solution Checker" for dialog as #checkerWnd
        gCheckerWndOpen = gTrue
        #checkerWnd.stepDelayTB "2"
        #checkerWnd "trapclose closeSolverWnd"
        #checkerWnd gFontString$
        wait
    end sub

    sub closeSolverWnd handle$
        timer 0
        close #handle$
        gCheckerWndOpen = gFalse
    end sub

    sub handleSolvedQuery handle$
        #checkerWnd.messagesTE "!cls"
        call upDateMsgs "Checking to see if your solution is correct ..."
        ' Check all the cells have been filled in
        if not(gridFilledIn()) then
            call upDateMsgs "Sorry your solution is not correct"
            exit sub
        end if

        ' Check all row and columns have 2 of each letter
        if not(gridLetterCounts()) then
            call upDateMsgs "Sorry your solution is not correct"
            exit sub
        end if

        ' Check clue conditions met
        if not(gridCluesMet()) then
            call upDateMsgs "Sorry your solution is not correct"
            exit sub
        end if

        call upDateMsgs "Congratulations! Solution is correct"
    end sub

    sub upDateMsgs newMsg$
        #checkerWnd.stepDelayTB "!contents? delay$"
        delay = int(val(delay$)) * 1000
        timer delay, [showNewMsg]
        wait
    [showNewMsg]
        timer 0
        #checkerWnd.messagesTE "!contents? curMsgs$"
        curMsgs$ = newMsg$ + chr$(13) + chr$(10) + curMsgs$
        #checkerWnd.messagesTE "!contents curMsgs$"
    end sub

    function setRowOrColNo$(rowColNo)
        if rowColNo <= 6 then
            setRowOrColNo$ = "Row " + str$(rowColNo)
        else
            setRowOrColNo$ = "Column " + str$(rowColNo - 6)
        end if
    end function

    ' rowColNo is an integer from 1 to 12 (1 to 6 is row, 7 to 12 is column)
    ' For given row or column load cell large labels into global vector string for processing
    sub extractVector rowColNo
        gVector$ = ""
        if rowColNo <= 6 then
            for col = 1 to 6
                gVector$ = gVector$ + largeLabels$(rowColNo, col)
            next col
        else
            for row = 1 to 6
                gVector$ = gVector$ + largeLabels$(row, rowColNo - 6)
            next row
        end if
    end sub

    function gridFilledIn()
        call upDateMsgs "Checking that all cells have a large label in them"
        gridFilledIn = gTrue
        for row = 1 to 6
            call extractVector row
            if not(vectorFilledIn(blanks)) then
                gridFilledIn = gFalse
                newMsg$ = "Row " + str$(row) + " has " + word$(gNumbers$, blanks + 1) +" vacant cells!"
                if instr(newMsg$, "one") then
                    newMsg$ = strReplace$(newMsg$, "vacant cells", "vacant cell")
                end if
                call upDateMsgs newMsg$
            end if
        next row
    end function

    function vectorFilledIn(byref blanks)
        vectorFilledIn = gTrue
        blanks = 0
        for index = 1 to 6
            if not(instr("ABC", mid$(gVector$, index, 1))) then
                vectorFilledIn = gFalse
                blanks = blanks + 1
            end if
        next index
    end function

    function gridLetterCounts()
        call upDateMsgs "Checking that each row and column has two A', B's and C's"
        gridLetterCounts = gTrue
        for rowColNo = 1 to 12
            call extractVector rowColNo
            for letter = asc("A") to asc("C")
                if not(vectorLetterCount(rowColNo, letter)) then
                    gridLetterCounts = gFalse
                end if
            next letter
        next rowColNo
    end function

    function vectorLetterCount(rowColNo, letter)
        vectorLetterCount = gTrue
        letterCount = 0
        for index = 1 to 6
            if mid$(gVector$, index, 1) = chr$(letter) then
                letterCount = letterCount + 1
            end if
        next index
        if letterCount <> 2 then
            rowColNo$ = setRowOrColNo$(rowColNo)
            newMsg$ = rowColNo$ + " has " + word$(gNumbers$, letterCount + 1) + " " + chr$(letter) + "'s"
            if instr(newMsg$, "one") then
                newMsg$ = mid$(newMsg$, 1, (len(newMsg$) - 2)) ' strip 's
            end if
            call upDateMsgs newMsg$
            vectorLetterCount = gFalse
        end if
    end function

    function gridCluesMet()
        call upDateMsgs "Checking that each row and column clue is satisfied"
        gridCluesMet = gTrue
        for clueNo = 1 to 12
            templateNo = val(puzzleInfo$(clueNo, 1))
            if templateNo <> 0 then
                select case templateNo
                    case 1
                        if not(vectorTemp1Met(clueNo)) then gridCluesMet = gFalse
                    case 2
                        if not(vectorTemp2Met(clueNo)) then gridCluesMet = gFalse
                    case 3
                        if not(vectorTemp3Met(clueNo)) then gridCluesMet = gFalse
                    case 4
                        if not(vectorTemp4Met(clueNo)) then gridCluesMet = gFalse
                    case 5
                        if not(vectorTemp5Met(clueNo)) then gridCluesMet = gFalse
                    case 6
                        if not(vectorTemp6Met(clueNo)) then gridCluesMet = gFalse
                end select
            end if
        next clueNo
    end function

    ' Checking template "The X's are in adjacent squares"
    function vectorTemp1Met(rowColNo)
        vectorTemp1Met = gFalse
        call extractVector rowColNo
        xValue$ = puzzleInfo$(rowColNo, 2)
        check$ = xValue$ + xValue$
        if instr(gVector$, check$) then vectorTemp1Met = gTrue
        if not(vectorTemp1Met) then
            rowColNo$ = setRowOrColNo$(rowColNo)
            newMsg$ = rowColNo$ + " does not meet clue condition"
            call upDateMsgs newMsg$
        end if
    end function

    ' Checking template "The X's are beyond the Y's"
    function vectorTemp2Met(rowColNo)
        vectorTemp2Met = gTrue
        call extractVector rowColNo
        xValue$ = puzzleInfo$(rowColNo, 2)
        yValue$ = puzzleInfo$(rowColNo, 3)
        firstX = instr(gVector$, xValue$)
        firstY = instr(gVector$, yValue$)
        secondY = instr(gVector$, yValue$, firstY + 1)
        if firstX < secondY then vectorTemp2Met = gFalse
        if not(vectorTemp2Met) then
            rowColNo$ = setRowOrColNo$(rowColNo)
            newMsg$ = rowColNo$ + " does not meet clue condition"
            call upDateMsgs newMsg$
        end if
    end function

    ' Checking template "The X's are between the Y's"
    function vectorTemp3Met(rowColNo)
        vectorTemp3Met = gTrue
        call extractVector rowColNo
        xValue$ = puzzleInfo$(rowColNo, 2)
        yValue$ = puzzleInfo$(rowColNo, 3)
        firstX = instr(gVector$, xValue$)
        secondX = instr(gVector$, xValue$, firstX + 1)
        firstY = instr(gVector$, yValue$)
        secondY = instr(gVector$, yValue$, firstY + 1)
        if (firstX < firstY) or (secondX > secondY) then vectorTemp3Met = gFalse
        if not(vectorTemp3Met) then
            rowColNo$ = setRowOrColNo$(rowColNo)
            newMsg$ = rowColNo$ + " does not meet clue condition"
            call upDateMsgs newMsg$
        end if
    end function

    ' Checking template "Each X is next to and beyond a Y"
    function vectorTemp4Met(rowColNo)
        vectorTemp4Met = gTrue
        call extractVector rowColNo
        xValue$ = puzzleInfo$(rowColNo, 2)
        yValue$ = puzzleInfo$(rowColNo, 3)
        firstX = instr(gVector$, xValue$)
        secondX = instr(gVector$, xValue$, firstX + 1)
        firstY = instr(gVector$, yValue$)
        secondY = instr(gVector$, yValue$, firstY + 1)
        if (firstX <> firstY + 1) or (secondX <> secondY + 1) then vectorTemp4Met = gFalse
        if not(vectorTemp4Met) then
            rowColNo$ = setRowOrColNo$(rowColNo)
            newMsg$ = rowColNo$ + " does not meet clue condition"
            call upDateMsgs newMsg$
        end if
    end function

    ' Checking template "No two adjacent squares contain the same letter"
    function vectorTemp5Met(rowColNo)
        vectorTemp5Met = gTrue
        call extractVector rowColNo
        for letterPos = 1 to 5
            if mid$(gVector$, letterPos, 1) = mid$(gVector$, letterPos + 1, 1) then vectorTemp5Met = gFalse
        next letterPos
        if not(vectorTemp5Met) then
            rowColNo$ = setRowOrColNo$(rowColNo)
            newMsg$ = rowColNo$ + " does not meet clue condition"
            call upDateMsgs newMsg$
        end if
    end function

    ' Checking template "Any 3 consecutive squares contain 3 different letters"
    function vectorTemp6Met(rowColNo)
        vectorTemp6Met = gTrue
        call extractVector rowColNo
        for letterPos = 1 to 4
            if (mid$(gVector$, letterPos, 1) = mid$(gVector$, letterPos + 1, 1)) or_
               (mid$(gVector$, letterPos, 1) = mid$(gVector$, letterPos + 2, 1)) or_
               (mid$(gVector$, letterPos + 1, 1) = mid$(gVector$, letterPos + 2, 1)) then vectorTemp6Met = gFalse
        next letterPos
        if not(vectorTemp6Met) then
            rowColNo$ = setRowOrColNo$(rowColNo)
            newMsg$ = rowColNo$ + " does not meet clue condition"
            call upDateMsgs newMsg$
        end if
    end function

    '--------------------------------------------------------------------------------------------------------------
    '
    '                   "States" Window
    '
    '--------------------------------------------------------------------------------------------------------------


    ' User can record up to 10 states. A state is a set of large and small button labels at
    ' a given point in solving the puzzle. This allows the user to try alternate paths looking
    ' for contradictions to eliminate possible labels for a given cell. So if user believes
    ' a cell can be A or B they first save state. Then set the cell to A and keep going. If they
    ' reach a contradiction then they know the cell must be B. So they load the saved state and set
    ' the cell to B and proceed from there.
    sub openStatesWnd
        WindowWidth = 410
        WindowHeight = 465
        UpperLeftX = gMainUpperLeftX + gMainWindowWidth
        UpperLeftY = gMainUpperLeftY
        groupbox #statesWnd.descriptionsGB, "States", 21, 16, 360, 290
        checkbox #statesWnd.state1, "1", stateCBselected, dummyHandler, 46, 46, 24, 20
        checkbox #statesWnd.state2, "2", stateCBselected, dummyHandler, 46, 71, 24, 20
        checkbox #statesWnd.state3, "3", stateCBselected, dummyHandler, 46, 96, 24, 20
        checkbox #statesWnd.state4, "4", stateCBselected, dummyHandler, 46, 121, 24, 20
        checkbox #statesWnd.state5, "5", stateCBselected, dummyHandler, 46, 146, 24, 20
        checkbox #statesWnd.state6, "6", stateCBselected, dummyHandler, 46, 171, 24, 20
        checkbox #statesWnd.state7, "7", stateCBselected, dummyHandler, 46, 196, 24, 20
        checkbox #statesWnd.state8, "8", stateCBselected, dummyHandler, 46, 221, 24, 20
        checkbox #statesWnd.state9, "9", stateCBselected, dummyHandler, 46, 246, 24, 20
        checkbox #statesWnd.state10, "10", stateCBselected, dummyHandler, 46, 271, 32, 20
        statictext #statesWnd.stateDesc1, "empty", 86, 48, 280, 20
        statictext #statesWnd.stateDesc2, "empty", 86, 73, 280, 20
        statictext #statesWnd.stateDesc3, "empty", 86, 98, 280, 20
        statictext #statesWnd.stateDesc4, "empty", 86, 123, 280, 20
        statictext #statesWnd.stateDesc5, "empty", 86, 148, 280, 20
        statictext #statesWnd.stateDesc6, "empty", 86, 173, 280, 20
        statictext #statesWnd.stateDesc7, "empty", 86, 198, 280, 20
        statictext #statesWnd.stateDesc8, "empty", 86, 223, 280, 20
        statictext #statesWnd.stateDesc9, "empty", 86, 248, 280, 20
        statictext #statesWnd.stateDesc10, "empty", 86, 273, 280, 20
        groupbox #statesWnd.actionsGB, "State Description", 21, 316, 360, 105
        textbox #statesWnd.DescriptionTB, 51, 341, 295, 25
        button #statesWnd.saveB, "Save", saveStateHandler, UL, 151, 376, 100, 25
        button #statesWnd.loadB, "Load", loadStateHandler, UL, 36, 376, 100, 25
        button #statesWnd.clearB, "Clear", clearStateHandler, UL, 266, 376, 100, 25
        open "States" for window_nf as #statesWnd
        gStatesWndOpen = gTrue
        #statesWnd gFontString$
        #statesWnd "trapclose closeStatesWnd"
        #statesWnd.state1, "set"
        stateChecked$(1) = "1"
        call refreshStates
        gGameChanged = gFalse
        scan
    end sub

    sub closeStatesWnd handle$
        close #handle$
        gStatesWndOpen = gFalse
    end sub

    sub refreshStates
        if gStatesWndOpen = gFalse then exit sub
        for state = 1 to 10
            ' Refresh state checkboxes
            handleCheckbox$ = "#statesWnd.state" + str$(state)
            if stateChecked$(state) = "1" then
                #handleCheckbox$ "set"
            else
                #handleCheckbox$ "reset"
            end if

            ' Refresh state descriptions
            handleStateDesc$ = "#statesWnd.stateDesc" + str$(state)
            #handleStateDesc$ stateDescriptions$(state)
            if stateChecked$(state) = "1" then
                #handleStateDesc$ "!font " + gFontString$ + " bold"
            else
                #handleStateDesc$ "!font " + gFontString$
            end if
        next state
        gGameChanged = gTrue
    end sub

    ' Called when 1 of the 10 state checkboxes is selected and resets remaining 9
    ' effectively giving a 10 option radio button
    sub stateCBselected handle$
        clickedCheckbox = val(mid$(handle$, 17))
        for index = 1 to 10
            if index <> clickedCheckbox then
                handle$ = "#statesWnd.state"; str$(index)
                #handle$, "reset"
                stateChecked$(index) = "0"
            else
                stateChecked$(index) = "1"
            end if
        next index
        call refreshStates
    end sub

    ' Identify which state is currently selected
    function getSelectedState()
        for index = 1 to 10
            handle$ = "#statesWnd.state"; str$(index)
            #handle$, "value? result$"
            if result$ = "set" then
                clickedCheckbox = index
                exit for
            end if
        next index
        getSelectedState = clickedCheckbox
    end function

    sub loadStateHandler handle$
        selectedState = getSelectedState()
        confirm "This will delete any current"_
                + chr$(13) + "large and small letters from the grid"_
                + chr$(13) + chr$(13) + "Do you still want to load from state "_
                + str$(selectedState) + "?"; answer$
        if answer$ = "no" then wait

        ' Load large and small button labels from state array
        for row = 1 to 6
            for col = 1 to 6
                largeLabels$(row, col) = stateLargeLabels$(selectedState, row * 7 + col)
                smallLabelsTriple$ = stateSmallLabels$(selectedState, row * 7 + col)
                smallLabels$(row, col * 6 + 2) = mid$(smallLabelsTriple$, 1, 1)
                smallLabels$(row, col * 6 + 3) = mid$(smallLabelsTriple$, 2, 1)
                smallLabels$(row, col * 6 + 4) = mid$(smallLabelsTriple$, 3, 1)
            next col
        next row
        call refreshBitmaps

        ' Load which clue conditions met checkboxes have been ticked
        for clueNo = 1 to 12
            clueConditionMet$(clueNo) = stateClueConditionMet$(selectedState, clueNo)
        next clueNo
        call refreshClues
    end sub

    sub saveStateHandler handle$
        selectedState = getSelectedState()
        confirm "This will replace any existing stored"_
                + chr$(13) + "large and small letters for the state"_
                + chr$(13) + chr$(13) + "Do you still want to save to state "_
                + str$(selectedState) + "?"; answer$
        if answer$ = "no" then wait
        ' Update selected state description
        handleStateDesc$ = "#statesWnd.stateDesc" + str$(selectedState)
        #statesWnd.DescriptionTB "!contents? newStateDesc$"
        #handleStateDesc$ newStateDesc$
        stateDescriptions$(selectedState) = newStateDesc$
        #statesWnd.DescriptionTB ""

        ' Save large and small button labels in state array
        for row = 1 to 6
            for col = 1 to 6
                stateLargeLabels$(selectedState, row * 7 + col) = largeLabels$(row, col)
                smallLabelsTriple$ = smallLabels$(row, col * 6 + 2) + smallLabels$(row, col * 6 + 3) +_
                                     smallLabels$(row, col * 6 + 4)
                stateSmallLabels$(selectedState, row * 7 + col) = smallLabelsTriple$
            next col
        next row

        ' Save which clue conditions met checkboxes have been ticked
        for clueNo = 1 to 12
            stateClueConditionMet$(selectedState, clueNo) = clueConditionMet$(clueNo)
        next clueNo
    end sub

    sub clearStateHandler handle$
        selectedState = getSelectedState()
        confirm "This will delete any existing stored"_
                + chr$(13) + "large and small letters for the state"_
                + chr$(13) + chr$(13) + "Do you still want to clear state "_
                + str$(selectedState) + "?"; answer$
        if answer$ = "no" then wait
        ' Update selected state description
        handleStateDesc$ = "#statesWnd.stateDesc" + str$(selectedState)
        #handleStateDesc$ "empty"
        stateDescriptions$(selectedState) = "empty"
        call refreshStates

        ' Clear large and small button labels in state array
        for row = 1 to 6
            for col = 1 to 6
                stateLargeLabels$(selectedState, row * 7 + col) = "?"
                stateSmallLabels$(selectedState, row * 7 + col) = "???"
            next col
        next row
    end sub

    sub dummyHandler handle$
    end sub

    sub createButtons
        bmpbutton #mainWnd.cell11, "bmp\6NE.bmp", handleLargeClick, UL, 0, 20
        bmpbutton #mainWnd.cell11.1, "bmp\1NE.bmp", handleSmallClick, UL, 0, 0
        bmpbutton #mainWnd.cell11.2, "bmp\2NE.bmp", handleSmallClick, UL, 20, 0
        bmpbutton #mainWnd.cell11.3, "bmp\3NE.bmp", handleSmallClick, UL, 40, 0
        bmpbutton #mainWnd.cell11.4, "bmp\4NE.bmp", handleSmallClick, UL, 60, 0
        bmpbutton #mainWnd.cell11.5, "bmp\5NE.bmp", handleSmallClick, UL, 80, 0
        bmpbutton #mainWnd.cell12, "bmp\6NE.bmp", handleLargeClick, UL, 100, 20
        bmpbutton #mainWnd.cell12.1, "bmp\1NE.bmp", handleSmallClick, UL, 100, 0
        bmpbutton #mainWnd.cell12.2, "bmp\2NE.bmp", handleSmallClick, UL, 120, 0
        bmpbutton #mainWnd.cell12.3, "bmp\3NE.bmp", handleSmallClick, UL, 140, 0
        bmpbutton #mainWnd.cell12.4, "bmp\4NE.bmp", handleSmallClick, UL, 160, 0
        bmpbutton #mainWnd.cell12.5, "bmp\5NE.bmp", handleSmallClick, UL, 180, 0
        bmpbutton #mainWnd.cell13, "bmp\6NE.bmp", handleLargeClick, UL, 200, 20
        bmpbutton #mainWnd.cell13.1, "bmp\1NE.bmp", handleSmallClick, UL, 200, 0
        bmpbutton #mainWnd.cell13.2, "bmp\2NE.bmp", handleSmallClick, UL, 220, 0
        bmpbutton #mainWnd.cell13.3, "bmp\3NE.bmp", handleSmallClick, UL, 240, 0
        bmpbutton #mainWnd.cell13.4, "bmp\4NE.bmp", handleSmallClick, UL, 260, 0
        bmpbutton #mainWnd.cell13.5, "bmp\5NE.bmp", handleSmallClick, UL, 280, 0
        bmpbutton #mainWnd.cell14, "bmp\6NE.bmp", handleLargeClick, UL, 300, 20
        bmpbutton #mainWnd.cell14.1, "bmp\1NE.bmp", handleSmallClick, UL, 300, 0
        bmpbutton #mainWnd.cell14.2, "bmp\2NE.bmp", handleSmallClick, UL, 320, 0
        bmpbutton #mainWnd.cell14.3, "bmp\3NE.bmp", handleSmallClick, UL, 340, 0
        bmpbutton #mainWnd.cell14.4, "bmp\4NE.bmp", handleSmallClick, UL, 360, 0
        bmpbutton #mainWnd.cell14.5, "bmp\5NE.bmp", handleSmallClick, UL, 380, 0
        bmpbutton #mainWnd.cell15, "bmp\6NE.bmp", handleLargeClick, UL, 400, 20
        bmpbutton #mainWnd.cell15.1, "bmp\1NE.bmp", handleSmallClick, UL, 400, 0
        bmpbutton #mainWnd.cell15.2, "bmp\2NE.bmp", handleSmallClick, UL, 420, 0
        bmpbutton #mainWnd.cell15.3, "bmp\3NE.bmp", handleSmallClick, UL, 440, 0
        bmpbutton #mainWnd.cell15.4, "bmp\4NE.bmp", handleSmallClick, UL, 460, 0
        bmpbutton #mainWnd.cell15.5, "bmp\5NE.bmp", handleSmallClick, UL, 480, 0
        bmpbutton #mainWnd.cell16, "bmp\6NE.bmp", handleLargeClick, UL, 500, 20
        bmpbutton #mainWnd.cell16.1, "bmp\1NE.bmp", handleSmallClick, UL, 500, 0
        bmpbutton #mainWnd.cell16.2, "bmp\2NE.bmp", handleSmallClick, UL, 520, 0
        bmpbutton #mainWnd.cell16.3, "bmp\3NE.bmp", handleSmallClick, UL, 540, 0
        bmpbutton #mainWnd.cell16.4, "bmp\4NE.bmp", handleSmallClick, UL, 560, 0
        bmpbutton #mainWnd.cell16.5, "bmp\5NE.bmp", handleSmallClick, UL, 580, 0
        bmpbutton #mainWnd.cell21, "bmp\6NE.bmp", handleLargeClick, UL, 0, 120
        bmpbutton #mainWnd.cell21.1, "bmp\1NE.bmp", handleSmallClick, UL, 0, 100
        bmpbutton #mainWnd.cell21.2, "bmp\2NE.bmp", handleSmallClick, UL, 20, 100
        bmpbutton #mainWnd.cell21.3, "bmp\3NE.bmp", handleSmallClick, UL, 40, 100
        bmpbutton #mainWnd.cell21.4, "bmp\4NE.bmp", handleSmallClick, UL, 60, 100
        bmpbutton #mainWnd.cell21.5, "bmp\5NE.bmp", handleSmallClick, UL, 80, 100
        bmpbutton #mainWnd.cell22, "bmp\6NE.bmp", handleLargeClick, UL, 100, 120
        bmpbutton #mainWnd.cell22.1, "bmp\1NE.bmp", handleSmallClick, UL, 100, 100
        bmpbutton #mainWnd.cell22.2, "bmp\2NE.bmp", handleSmallClick, UL, 120, 100
        bmpbutton #mainWnd.cell22.3, "bmp\3NE.bmp", handleSmallClick, UL, 140, 100
        bmpbutton #mainWnd.cell22.4, "bmp\4NE.bmp", handleSmallClick, UL, 160, 100
        bmpbutton #mainWnd.cell22.5, "bmp\5NE.bmp", handleSmallClick, UL, 180, 100
        bmpbutton #mainWnd.cell23, "bmp\6NE.bmp", handleLargeClick, UL, 200, 120
        bmpbutton #mainWnd.cell23.1, "bmp\1NE.bmp", handleSmallClick, UL, 200, 100
        bmpbutton #mainWnd.cell23.2, "bmp\2NE.bmp", handleSmallClick, UL, 220, 100
        bmpbutton #mainWnd.cell23.3, "bmp\3NE.bmp", handleSmallClick, UL, 240, 100
        bmpbutton #mainWnd.cell23.4, "bmp\4NE.bmp", handleSmallClick, UL, 260, 100
        bmpbutton #mainWnd.cell23.5, "bmp\5NE.bmp", handleSmallClick, UL, 280, 100
        bmpbutton #mainWnd.cell24, "bmp\6NE.bmp", handleLargeClick, UL, 300, 120
        bmpbutton #mainWnd.cell24.1, "bmp\1NE.bmp", handleSmallClick, UL, 300, 100
        bmpbutton #mainWnd.cell24.2, "bmp\2NE.bmp", handleSmallClick, UL, 320, 100
        bmpbutton #mainWnd.cell24.3, "bmp\3NE.bmp", handleSmallClick, UL, 340, 100
        bmpbutton #mainWnd.cell24.4, "bmp\4NE.bmp", handleSmallClick, UL, 360, 100
        bmpbutton #mainWnd.cell24.5, "bmp\5NE.bmp", handleSmallClick, UL, 380, 100
        bmpbutton #mainWnd.cell25, "bmp\6NE.bmp", handleLargeClick, UL, 400, 120
        bmpbutton #mainWnd.cell25.1, "bmp\1NE.bmp", handleSmallClick, UL, 400, 100
        bmpbutton #mainWnd.cell25.2, "bmp\2NE.bmp", handleSmallClick, UL, 420, 100
        bmpbutton #mainWnd.cell25.3, "bmp\3NE.bmp", handleSmallClick, UL, 440, 100
        bmpbutton #mainWnd.cell25.4, "bmp\4NE.bmp", handleSmallClick, UL, 460, 100
        bmpbutton #mainWnd.cell25.5, "bmp\5NE.bmp", handleSmallClick, UL, 480, 100
        bmpbutton #mainWnd.cell26, "bmp\6NE.bmp", handleLargeClick, UL, 500, 120
        bmpbutton #mainWnd.cell26.1, "bmp\1NE.bmp", handleSmallClick, UL, 500, 100
        bmpbutton #mainWnd.cell26.2, "bmp\2NE.bmp", handleSmallClick, UL, 520, 100
        bmpbutton #mainWnd.cell26.3, "bmp\3NE.bmp", handleSmallClick, UL, 540, 100
        bmpbutton #mainWnd.cell26.4, "bmp\4NE.bmp", handleSmallClick, UL, 560, 100
        bmpbutton #mainWnd.cell26.5, "bmp\5NE.bmp", handleSmallClick, UL, 580, 100
        bmpbutton #mainWnd.cell31, "bmp\6NE.bmp", handleLargeClick, UL, 0, 220
        bmpbutton #mainWnd.cell31.1, "bmp\1NE.bmp", handleSmallClick, UL, 0, 200
        bmpbutton #mainWnd.cell31.2, "bmp\2NE.bmp", handleSmallClick, UL, 20, 200
        bmpbutton #mainWnd.cell31.3, "bmp\3NE.bmp", handleSmallClick, UL, 40, 200
        bmpbutton #mainWnd.cell31.4, "bmp\4NE.bmp", handleSmallClick, UL, 60, 200
        bmpbutton #mainWnd.cell31.5, "bmp\5NE.bmp", handleSmallClick, UL, 80, 200
        bmpbutton #mainWnd.cell32, "bmp\6NE.bmp", handleLargeClick, UL, 100, 220
        bmpbutton #mainWnd.cell32.1, "bmp\1NE.bmp", handleSmallClick, UL, 100, 200
        bmpbutton #mainWnd.cell32.2, "bmp\2NE.bmp", handleSmallClick, UL, 120, 200
        bmpbutton #mainWnd.cell32.3, "bmp\3NE.bmp", handleSmallClick, UL, 140, 200
        bmpbutton #mainWnd.cell32.4, "bmp\4NE.bmp", handleSmallClick, UL, 160, 200
        bmpbutton #mainWnd.cell32.5, "bmp\5NE.bmp", handleSmallClick, UL, 180, 200
        bmpbutton #mainWnd.cell33, "bmp\6NE.bmp", handleLargeClick, UL, 200, 220
        bmpbutton #mainWnd.cell33.1, "bmp\1NE.bmp", handleSmallClick, UL, 200, 200
        bmpbutton #mainWnd.cell33.2, "bmp\2NE.bmp", handleSmallClick, UL, 220, 200
        bmpbutton #mainWnd.cell33.3, "bmp\3NE.bmp", handleSmallClick, UL, 240, 200
        bmpbutton #mainWnd.cell33.4, "bmp\4NE.bmp", handleSmallClick, UL, 260, 200
        bmpbutton #mainWnd.cell33.5, "bmp\5NE.bmp", handleSmallClick, UL, 280, 200
        bmpbutton #mainWnd.cell34, "bmp\6NE.bmp", handleLargeClick, UL, 300, 220
        bmpbutton #mainWnd.cell34.1, "bmp\1NE.bmp", handleSmallClick, UL, 300, 200
        bmpbutton #mainWnd.cell34.2, "bmp\2NE.bmp", handleSmallClick, UL, 320, 200
        bmpbutton #mainWnd.cell34.3, "bmp\3NE.bmp", handleSmallClick, UL, 340, 200
        bmpbutton #mainWnd.cell34.4, "bmp\4NE.bmp", handleSmallClick, UL, 360, 200
        bmpbutton #mainWnd.cell34.5, "bmp\5NE.bmp", handleSmallClick, UL, 380, 200
        bmpbutton #mainWnd.cell35, "bmp\6NE.bmp", handleLargeClick, UL, 400, 220
        bmpbutton #mainWnd.cell35.1, "bmp\1NE.bmp", handleSmallClick, UL, 400, 200
        bmpbutton #mainWnd.cell35.2, "bmp\2NE.bmp", handleSmallClick, UL, 420, 200
        bmpbutton #mainWnd.cell35.3, "bmp\3NE.bmp", handleSmallClick, UL, 440, 200
        bmpbutton #mainWnd.cell35.4, "bmp\4NE.bmp", handleSmallClick, UL, 460, 200
        bmpbutton #mainWnd.cell35.5, "bmp\5NE.bmp", handleSmallClick, UL, 480, 200
        bmpbutton #mainWnd.cell36, "bmp\6NE.bmp", handleLargeClick, UL, 500, 220
        bmpbutton #mainWnd.cell36.1, "bmp\1NE.bmp", handleSmallClick, UL, 500, 200
        bmpbutton #mainWnd.cell36.2, "bmp\2NE.bmp", handleSmallClick, UL, 520, 200
        bmpbutton #mainWnd.cell36.3, "bmp\3NE.bmp", handleSmallClick, UL, 540, 200
        bmpbutton #mainWnd.cell36.4, "bmp\4NE.bmp", handleSmallClick, UL, 560, 200
        bmpbutton #mainWnd.cell36.5, "bmp\5NE.bmp", handleSmallClick, UL, 580, 200
        bmpbutton #mainWnd.cell41, "bmp\6NE.bmp", handleLargeClick, UL, 0, 320
        bmpbutton #mainWnd.cell41.1, "bmp\1NE.bmp", handleSmallClick, UL, 0, 300
        bmpbutton #mainWnd.cell41.2, "bmp\2NE.bmp", handleSmallClick, UL, 20, 300
        bmpbutton #mainWnd.cell41.3, "bmp\3NE.bmp", handleSmallClick, UL, 40, 300
        bmpbutton #mainWnd.cell41.4, "bmp\4NE.bmp", handleSmallClick, UL, 60, 300
        bmpbutton #mainWnd.cell41.5, "bmp\5NE.bmp", handleSmallClick, UL, 80, 300
        bmpbutton #mainWnd.cell42, "bmp\6NE.bmp", handleLargeClick, UL, 100, 320
        bmpbutton #mainWnd.cell42.1, "bmp\1NE.bmp", handleSmallClick, UL, 100, 300
        bmpbutton #mainWnd.cell42.2, "bmp\2NE.bmp", handleSmallClick, UL, 120, 300
        bmpbutton #mainWnd.cell42.3, "bmp\3NE.bmp", handleSmallClick, UL, 140, 300
        bmpbutton #mainWnd.cell42.4, "bmp\4NE.bmp", handleSmallClick, UL, 160, 300
        bmpbutton #mainWnd.cell42.5, "bmp\5NE.bmp", handleSmallClick, UL, 180, 300
        bmpbutton #mainWnd.cell43, "bmp\6NE.bmp", handleLargeClick, UL, 200, 320
        bmpbutton #mainWnd.cell43.1, "bmp\1NE.bmp", handleSmallClick, UL, 200, 300
        bmpbutton #mainWnd.cell43.2, "bmp\2NE.bmp", handleSmallClick, UL, 220, 300
        bmpbutton #mainWnd.cell43.3, "bmp\3NE.bmp", handleSmallClick, UL, 240, 300
        bmpbutton #mainWnd.cell43.4, "bmp\4NE.bmp", handleSmallClick, UL, 260, 300
        bmpbutton #mainWnd.cell43.5, "bmp\5NE.bmp", handleSmallClick, UL, 280, 300
        bmpbutton #mainWnd.cell44, "bmp\6NE.bmp", handleLargeClick, UL, 300, 320
        bmpbutton #mainWnd.cell44.1, "bmp\1NE.bmp", handleSmallClick, UL, 300, 300
        bmpbutton #mainWnd.cell44.2, "bmp\2NE.bmp", handleSmallClick, UL, 320, 300
        bmpbutton #mainWnd.cell44.3, "bmp\3NE.bmp", handleSmallClick, UL, 340, 300
        bmpbutton #mainWnd.cell44.4, "bmp\4NE.bmp", handleSmallClick, UL, 360, 300
        bmpbutton #mainWnd.cell44.5, "bmp\5NE.bmp", handleSmallClick, UL, 380, 300
        bmpbutton #mainWnd.cell45, "bmp\6NE.bmp", handleLargeClick, UL, 400, 320
        bmpbutton #mainWnd.cell45.1, "bmp\1NE.bmp", handleSmallClick, UL, 400, 300
        bmpbutton #mainWnd.cell45.2, "bmp\2NE.bmp", handleSmallClick, UL, 420, 300
        bmpbutton #mainWnd.cell45.3, "bmp\3NE.bmp", handleSmallClick, UL, 440, 300
        bmpbutton #mainWnd.cell45.4, "bmp\4NE.bmp", handleSmallClick, UL, 460, 300
        bmpbutton #mainWnd.cell45.5, "bmp\5NE.bmp", handleSmallClick, UL, 480, 300
        bmpbutton #mainWnd.cell46, "bmp\6NE.bmp", handleLargeClick, UL, 500, 320
        bmpbutton #mainWnd.cell46.1, "bmp\1NE.bmp", handleSmallClick, UL, 500, 300
        bmpbutton #mainWnd.cell46.2, "bmp\2NE.bmp", handleSmallClick, UL, 520, 300
        bmpbutton #mainWnd.cell46.3, "bmp\3NE.bmp", handleSmallClick, UL, 540, 300
        bmpbutton #mainWnd.cell46.4, "bmp\4NE.bmp", handleSmallClick, UL, 560, 300
        bmpbutton #mainWnd.cell46.5, "bmp\5NE.bmp", handleSmallClick, UL, 580, 300
        bmpbutton #mainWnd.cell51, "bmp\6NE.bmp", handleLargeClick, UL, 0, 420
        bmpbutton #mainWnd.cell51.1, "bmp\1NE.bmp", handleSmallClick, UL, 0, 400
        bmpbutton #mainWnd.cell51.2, "bmp\2NE.bmp", handleSmallClick, UL, 20, 400
        bmpbutton #mainWnd.cell51.3, "bmp\3NE.bmp", handleSmallClick, UL, 40, 400
        bmpbutton #mainWnd.cell51.4, "bmp\4NE.bmp", handleSmallClick, UL, 60, 400
        bmpbutton #mainWnd.cell51.5, "bmp\5NE.bmp", handleSmallClick, UL, 80, 400
        bmpbutton #mainWnd.cell52, "bmp\6NE.bmp", handleLargeClick, UL, 100, 420
        bmpbutton #mainWnd.cell52.1, "bmp\1NE.bmp", handleSmallClick, UL, 100, 400
        bmpbutton #mainWnd.cell52.2, "bmp\2NE.bmp", handleSmallClick, UL, 120, 400
        bmpbutton #mainWnd.cell52.3, "bmp\3NE.bmp", handleSmallClick, UL, 140, 400
        bmpbutton #mainWnd.cell52.4, "bmp\4NE.bmp", handleSmallClick, UL, 160, 400
        bmpbutton #mainWnd.cell52.5, "bmp\5NE.bmp", handleSmallClick, UL, 180, 400
        bmpbutton #mainWnd.cell53, "bmp\6NE.bmp", handleLargeClick, UL, 200, 420
        bmpbutton #mainWnd.cell53.1, "bmp\1NE.bmp", handleSmallClick, UL, 200, 400
        bmpbutton #mainWnd.cell53.2, "bmp\2NE.bmp", handleSmallClick, UL, 220, 400
        bmpbutton #mainWnd.cell53.3, "bmp\3NE.bmp", handleSmallClick, UL, 240, 400
        bmpbutton #mainWnd.cell53.4, "bmp\4NE.bmp", handleSmallClick, UL, 260, 400
        bmpbutton #mainWnd.cell53.5, "bmp\5NE.bmp", handleSmallClick, UL, 280, 400
        bmpbutton #mainWnd.cell54, "bmp\6NE.bmp", handleLargeClick, UL, 300, 420
        bmpbutton #mainWnd.cell54.1, "bmp\1NE.bmp", handleSmallClick, UL, 300, 400
        bmpbutton #mainWnd.cell54.2, "bmp\2NE.bmp", handleSmallClick, UL, 320, 400
        bmpbutton #mainWnd.cell54.3, "bmp\3NE.bmp", handleSmallClick, UL, 340, 400
        bmpbutton #mainWnd.cell54.4, "bmp\4NE.bmp", handleSmallClick, UL, 360, 400
        bmpbutton #mainWnd.cell54.5, "bmp\5NE.bmp", handleSmallClick, UL, 380, 400
        bmpbutton #mainWnd.cell55, "bmp\6NE.bmp", handleLargeClick, UL, 400, 420
        bmpbutton #mainWnd.cell55.1, "bmp\1NE.bmp", handleSmallClick, UL, 400, 400
        bmpbutton #mainWnd.cell55.2, "bmp\2NE.bmp", handleSmallClick, UL, 420, 400
        bmpbutton #mainWnd.cell55.3, "bmp\3NE.bmp", handleSmallClick, UL, 440, 400
        bmpbutton #mainWnd.cell55.4, "bmp\4NE.bmp", handleSmallClick, UL, 460, 400
        bmpbutton #mainWnd.cell55.5, "bmp\5NE.bmp", handleSmallClick, UL, 480, 400
        bmpbutton #mainWnd.cell56, "bmp\6NE.bmp", handleLargeClick, UL, 500, 420
        bmpbutton #mainWnd.cell56.1, "bmp\1NE.bmp", handleSmallClick, UL, 500, 400
        bmpbutton #mainWnd.cell56.2, "bmp\2NE.bmp", handleSmallClick, UL, 520, 400
        bmpbutton #mainWnd.cell56.3, "bmp\3NE.bmp", handleSmallClick, UL, 540, 400
        bmpbutton #mainWnd.cell56.4, "bmp\4NE.bmp", handleSmallClick, UL, 560, 400
        bmpbutton #mainWnd.cell56.5, "bmp\5NE.bmp", handleSmallClick, UL, 580, 400
        bmpbutton #mainWnd.cell61, "bmp\6NE.bmp", handleLargeClick, UL, 0, 520
        bmpbutton #mainWnd.cell61.1, "bmp\1NE.bmp", handleSmallClick, UL, 0, 500
        bmpbutton #mainWnd.cell61.2, "bmp\2NE.bmp", handleSmallClick, UL, 20, 500
        bmpbutton #mainWnd.cell61.3, "bmp\3NE.bmp", handleSmallClick, UL, 40, 500
        bmpbutton #mainWnd.cell61.4, "bmp\4NE.bmp", handleSmallClick, UL, 60, 500
        bmpbutton #mainWnd.cell61.5, "bmp\5NE.bmp", handleSmallClick, UL, 80, 500
        bmpbutton #mainWnd.cell62, "bmp\6NE.bmp", handleLargeClick, UL, 100, 520
        bmpbutton #mainWnd.cell62.1, "bmp\1NE.bmp", handleSmallClick, UL, 100, 500
        bmpbutton #mainWnd.cell62.2, "bmp\2NE.bmp", handleSmallClick, UL, 120, 500
        bmpbutton #mainWnd.cell62.3, "bmp\3NE.bmp", handleSmallClick, UL, 140, 500
        bmpbutton #mainWnd.cell62.4, "bmp\4NE.bmp", handleSmallClick, UL, 160, 500
        bmpbutton #mainWnd.cell62.5, "bmp\5NE.bmp", handleSmallClick, UL, 180, 500
        bmpbutton #mainWnd.cell63, "bmp\6NE.bmp", handleLargeClick, UL, 200, 520
        bmpbutton #mainWnd.cell63.1, "bmp\1NE.bmp", handleSmallClick, UL, 200, 500
        bmpbutton #mainWnd.cell63.2, "bmp\2NE.bmp", handleSmallClick, UL, 220, 500
        bmpbutton #mainWnd.cell63.3, "bmp\3NE.bmp", handleSmallClick, UL, 240, 500
        bmpbutton #mainWnd.cell63.4, "bmp\4NE.bmp", handleSmallClick, UL, 260, 500
        bmpbutton #mainWnd.cell63.5, "bmp\5NE.bmp", handleSmallClick, UL, 280, 500
        bmpbutton #mainWnd.cell64, "bmp\6NE.bmp", handleLargeClick, UL, 300, 520
        bmpbutton #mainWnd.cell64.1, "bmp\1NE.bmp", handleSmallClick, UL, 300, 500
        bmpbutton #mainWnd.cell64.2, "bmp\2NE.bmp", handleSmallClick, UL, 320, 500
        bmpbutton #mainWnd.cell64.3, "bmp\3NE.bmp", handleSmallClick, UL, 340, 500
        bmpbutton #mainWnd.cell64.4, "bmp\4NE.bmp", handleSmallClick, UL, 360, 500
        bmpbutton #mainWnd.cell64.5, "bmp\5NE.bmp", handleSmallClick, UL, 380, 500
        bmpbutton #mainWnd.cell65, "bmp\6NE.bmp", handleLargeClick, UL, 400, 520
        bmpbutton #mainWnd.cell65.1, "bmp\1NE.bmp", handleSmallClick, UL, 400, 500
        bmpbutton #mainWnd.cell65.2, "bmp\2NE.bmp", handleSmallClick, UL, 420, 500
        bmpbutton #mainWnd.cell65.3, "bmp\3NE.bmp", handleSmallClick, UL, 440, 500
        bmpbutton #mainWnd.cell65.4, "bmp\4NE.bmp", handleSmallClick, UL, 460, 500
        bmpbutton #mainWnd.cell65.5, "bmp\5NE.bmp", handleSmallClick, UL, 480, 500
        bmpbutton #mainWnd.cell66, "bmp\6NE.bmp", handleLargeClick, UL, 500, 520
        bmpbutton #mainWnd.cell66.1, "bmp\1NE.bmp", handleSmallClick, UL, 500, 500
        bmpbutton #mainWnd.cell66.2, "bmp\2NE.bmp", handleSmallClick, UL, 520, 500
        bmpbutton #mainWnd.cell66.3, "bmp\3NE.bmp", handleSmallClick, UL, 540, 500
        bmpbutton #mainWnd.cell66.4, "bmp\4NE.bmp", handleSmallClick, UL, 560, 500
        bmpbutton #mainWnd.cell66.5, "bmp\5NE.bmp", handleSmallClick, UL, 580, 500
    end sub

    ' Library code from external sources

    ' From QB64 and found for me by + thanks
    Function strReplace$(s$, replace$, new$) 'case sensitive  QB64 2020-07-28 version
        If Len(s$) = 0 Or Len(replace$) = 0 Then
            strReplace$ = s$: Exit Function
        Else
            LR = Len(replace$): lNew = Len(new$)
        End If
        sCopy$ = s$ ' otherwise s$ would get changed in qb64
        p = InStr(sCopy$, replace$)
        While p
            sCopy$ = Mid$(sCopy$, 1, p - 1) + new$ + Mid$(sCopy$, p + LR)
            p = InStr(sCopy$, replace$, p + lNew) ' InStr is differnt in JB
        Wend
        strReplace$ = sCopy$
    End Function

    ' From "jb20help" html
    function fileExists(path$, filename$)
        files path$, filename$, info$()
        fileExists = val(info$(0, 0)) 'non zero is true
    end function
