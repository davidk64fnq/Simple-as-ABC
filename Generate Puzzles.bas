    global gRowLetters$, gVector$, gPuzzle$
    dim largeLabels$(7,7)   ' A, B, C
    dim colLetters$(7)
    dim combos$(7)
    combos$(1) = "AB"
    combos$(2) = "BA"
    combos$(3) = "AC"
    combos$(4) = "CA"
    combos$(5) = "BC"
    combos$(6) = "CB"
    dim clues$(13)

    for count = 1 to 500000
        do
        loop while not(createRandomArray())
        gPuzzle$ = ""
        for clueNo = 1 to 12
            clues$(clueNo) = ""
        next clueNo
        call setRules
        if getClueCount() = 12 then
            ' Build puzzle
            for clueNo = 1 to 12
                if clues$(clueNo) <> "" then gPuzzle$ = gPuzzle$ + clues$(clueNo)
            next clueNo
            ' Remove trailing comma
            if gPuzzle$ <> "" then gPuzzle$ = left$(gPuzzle$, len(gPuzzle$) - 1)
'            call printArray
            print gPuzzle$
        end if
    next count
    print "finished"

    sub setRules
        call tryEachTemplateOnce
        if getClueCount() = 6 then call tryEachTemplateOnce
    end sub

    sub tryEachTemplateOnce
        start = int(rnd(1) * 12) + 1     ' Start from random row/column
        for index = start to start + 11
            rowColNo = index mod 12 + 1
            if vectorTemp6Met(rowColNo) then exit for
        next index
        for index = start to start + 11
            rowColNo = index mod 12 + 1
            if vectorTemp4Met(rowColNo) then exit for
        next index
        for index = start to start + 11
            rowColNo = index mod 12 + 1
            if vectorTemp3Met(rowColNo) then exit for
        next index
        for index = start to start + 11
            rowColNo = index mod 12 + 1
            if vectorTemp2Met(rowColNo) then exit for
        next index
        for index = start to start + 11
            rowColNo = index mod 12 + 1
            if vectorTemp5Met(rowColNo) then exit for
        next index
        for index = start to start + 11
            rowColNo = index mod 12 + 1
            if vectorTemp1Met(rowColNo) then exit for
        next index
    end sub

    function getClueCount()
        for clueNo = 1 to 12
            if clues$(clueNo) <> "" then getClueCount = getClueCount + 1
        next clueNo
    end function

    ' Checking template "The X's are in adjacent squares"
    function vectorTemp1Met(rowColNo)
        vectorTemp1Met = 0
        call extractVector rowColNo
        start = int(rnd(1) * 3) + 1     ' Spread clues across the 3 letters
        for combo = start to start + 2
            comboIndex = combo mod 3 + 1
            xValue$ = chr$(asc("A") + comboIndex - 1)
            check$ = xValue$ + xValue$
            if instr(gVector$, check$) and (clues$(rowColNo) = "") then
                clues$(rowColNo) = str$(rowColNo) + ".1" + lower$(xValue$) + ","
                vectorTemp1Met = 1
            end if
        next combo
    end function

    ' Checking template "The X's are beyond the Y's"
    function vectorTemp2Met(rowColNo)
        vectorTemp2Met = 0
        call extractVector rowColNo
        start = int(rnd(1) * 6) + 1     ' Spread clues across the 6 two letter pair combinations
        for combo = start to start + 5
            comboIndex = combo mod 6 + 1
            xValue$ = left$(combos$(comboIndex), 1)
            yValue$ = right$(combos$(comboIndex), 1)
            firstX = instr(gVector$, xValue$)
            firstY = instr(gVector$, yValue$)
            secondY = instr(gVector$, yValue$, firstY + 1)
            if (firstX > secondY) and (clues$(rowColNo) = "") then
                clues$(rowColNo) = str$(rowColNo) + ".2" + lower$(xValue$) + lower$(yValue$) + ","
                vectorTemp2Met = 1
            end if
        next combo
    end function

    ' Checking template "The X's are between the Y's"
    function vectorTemp3Met(rowColNo)
        vectorTemp3Met = 0
        call extractVector rowColNo
        start = int(rnd(1) * 6) + 1     ' Spread clues across the 6 two letter pair combinations
        for combo = start to start + 5
            comboIndex = combo mod 6 + 1
            xValue$ = left$(combos$(comboIndex), 1)
            yValue$ = right$(combos$(comboIndex), 1)
            firstX = instr(gVector$, xValue$)
            secondX = instr(gVector$, xValue$, firstX + 1)
            firstY = instr(gVector$, yValue$)
            secondY = instr(gVector$, yValue$, firstY + 1)
            if (firstX > firstY) and (secondX < secondY) and (clues$(rowColNo) = "") then
                clues$(rowColNo) = str$(rowColNo) + ".3" + lower$(xValue$) + lower$(yValue$) + ","
                vectorTemp3Met = 1
            end if
        next combo
    end function

    ' Checking template "Each X is next to and beyond a Y"
    function vectorTemp4Met(rowColNo)
        vectorTemp4Met = 0
        call extractVector rowColNo
        start = int(rnd(1) * 6) + 1     ' Spread clues across the 6 two letter pair combinations
        for combo = start to start + 5
            comboIndex = combo mod 6 + 1
            xValue$ = left$(combos$(comboIndex), 1)
            yValue$ = right$(combos$(comboIndex), 1)
            firstX = instr(gVector$, xValue$)
            secondX = instr(gVector$, xValue$, firstX + 1)
            firstY = instr(gVector$, yValue$)
            secondY = instr(gVector$, yValue$, firstY + 1)
            if (firstX = firstY + 1) and (secondX = secondY + 1) and (clues$(rowColNo) = "") then
                clues$(rowColNo) = str$(rowColNo) + ".4" + lower$(xValue$) + lower$(yValue$) + ","
                vectorTemp4Met = 1
            end if
        next combo
    end function

    ' Checking template "No two adjacent squares contain the same letter"
    function vectorTemp5Met(rowColNo)
        vectorTemp5Met = 1
        call extractVector rowColNo
        for letterPos = 1 to 5
            first$ = mid$(gVector$, letterPos, 1)
            second$ = mid$(gVector$, letterPos + 1, 1)
            if first$ = second$ then vectorTemp5Met = 0
        next letterPos
        if vectorTemp5Met and (clues$(rowColNo) = "") then clues$(rowColNo) = str$(rowColNo) + ".5"  + ","
    end function


    ' Checking template "Any 3 consecutive squares contain 3 different letters"
    function vectorTemp6Met(rowColNo)
        vectorTemp6Met = 1
        call extractVector rowColNo
        for letterPos = 1 to 4
            first$ = mid$(gVector$, letterPos, 1)
            second$ = mid$(gVector$, letterPos + 1, 1)
            third$ = mid$(gVector$, letterPos + 2, 1)
            if (first$ = second$) or (first$ = third$) or (second$ = third$) then vectorTemp6Met = 0
        next letterPos
        if vectorTemp6Met and (clues$(rowColNo) = "") then clues$(rowColNo) = str$(rowColNo) + ".6,"
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

    sub initArrays
        for row = 1 to 6
            for col = 1 to 6
                largeLabels$(row, col) = ""
            next col
        next row
        for col = 1 to 6
            colLetters$(col) = "AABBCC"
        next col
    end sub

    function createRandomArray()
        call initArrays
        createRandomArray = 1
        for row = 1 to 6
            gRowLetters$ = "AABBCC"
            for col = 1 to 6
                randomLetter$ = getRandomLetter$(col)
                if randomLetter$ <> "" then
                    largeLabels$(row, col) = randomLetter$
                else
                    largeLabels$(row, col) = "?"
                    createRandomArray = 0
                    exit function
                end if
            next col
        next row
    end function

    ' Attempt to find a random letter from gRowLetters in colLetters(colIndex)
    function getRandomLetter$(colIndex)
        getRandomLetter$ = ""
        rowStringLen = len(gRowLetters$)
        colStringLen = len(colLetters$(colIndex))
        do
            ' Get random letter from gRowLetters$
            rowPos = int(rnd(1) * rowStringLen) + 1
            rowLetter$ = mid$(gRowLetters$, rowPos, 1)

            ' See if rowLetter$ is available for use in colLetters$
            colPos = instr(colLetters$(colIndex), rowLetter$)
            if colPos then
                ' Remove rowLetter$ from colLetters$
                colLetters$(colIndex) = left$(colLetters$(colIndex), colPos - 1)_
                                      + right$(colLetters$(colIndex), colStringLen - colPos)

                ' Remove rowLetter$ from gRowLetters$
                gRowLetters$ = left$(gRowLetters$, rowPos - 1)_
                            + right$(gRowLetters$, rowStringLen - rowPos)

                getRandomLetter$ = rowLetter$
            end if
            loopCount = loopCount + 1
        loop while not(colPos) and loopCount < 100
    end function

    sub printArray
        print ""
        for row = 1 to 6
            for col = 1 to 6
                print largeLabels$(row, col);
            next col
            print ""
        next row
    end sub
