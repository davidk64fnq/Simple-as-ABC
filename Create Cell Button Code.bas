    gridSize = 20

    print "    sub createButtons"
    for row = 1 to 6
        cellPosLargeY = (row - 1) * gridSize * 5 + gridSize
        for col = 1 to 6
            buttonHandle$ = "#mainWnd.cell" + str$(row) + str$(col)
            cellPosLargeX = (col - 1) * gridSize * 5
            print "        bmpbutton " + buttonHandle$ + ", " + chr$(34) + "bmp\6NE.bmp" + chr$(34);
            print ", handleLargeClick, UL, " + str$(cellPosLargeX) + ", " + str$(cellPosLargeY)
            cellPosSmallY = (row - 1) * gridSize * 5
            for small = 1 to 5
                cellPosSmallX = (col - 1) * gridSize * 5 + (small - 1) * gridSize
                buttonHandle$ = "#mainWnd.cell" + str$(row) + str$(col) + "." + str$(small)
                bmpName$ = "bmp\" + str$(small) + "NE.bmp"
                print "        bmpbutton " + buttonHandle$ + ", " + chr$(34) + bmpName$ + chr$(34);
                print ", handleSmallClick, UL, " + str$(cellPosSmallX) + ", " + str$(cellPosSmallY)
            next small
        next col
    next row
    print "    end sub"
