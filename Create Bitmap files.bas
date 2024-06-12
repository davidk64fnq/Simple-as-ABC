    global gridSize, borderInset, boxInset
    gridSize = 20               ' Each big letter box is made up of 6 buttons
                                ' A row of 5 1 x 1 gridSize buttons across top
                                ' Below that a 4 deep 5 wide gridsize button
                                ' Top 5 buttons named 1 to 5, big button below named 6
    borderInset = 2             ' The 6 buttons have a border drawn on appropriate sides
                                ' to give appearance of one square
    boxInset = 6                ' The 2nd to 4th small buttons have a box to show they are
                                ' clickable to record letter possibles for main big letter

    call createBitmaps

    sub createBitmaps
        open "Drawing bitmaps" for graphics_nsb_nf as #gr
        #gr "down"
        call createNEbmpFiles "white"
        call createHEbmpFiles "buttonface"
        for index = asc("A") to asc("C")
            call createNltrBmpFile chr$(index), "white"
            call createHltrBmpFile chr$(index), "buttonface"
        next index
        wait
        close #gr
    end sub

    ' Naming convention for bitmaps is
    '   First character: 1 to 6 showing button position
    '   Second character: N = unhighlighted H = highlighted
    '   Third character: E = no letter A, B, C = the respective letter
    sub createNEbmpFiles bgColor$

        #gr "cls"
        call draw1BmpLines bgColor$, gridSize
        #gr "getbmp 1NE 0 0 "; str$(gridSize); " "; str$(gridSize);
        bmpsave "1NE", "bmp/1NE.bmp"

        #gr "cls"
        call draw2BmpLines bgColor$, gridSize
        #gr "getbmp 2NE 0 0 "; str$(gridSize); " "; str$(gridSize);
        bmpsave "2NE", "bmp/2NE.bmp"

        #gr "cls"
        call draw3BmpLines bgColor$, gridSize
        #gr "getbmp 3NE 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "3NE", "bmp/3NE.bmp"

        #gr "cls"
        call draw4BmpLines bgColor$, gridSize
        #gr "getbmp 4NE 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "4NE", "bmp/4NE.bmp"

        #gr "cls"
        call draw5BmpLines bgColor$, gridSize
        #gr "getbmp 5NE 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "5NE", "bmp/5NE.bmp"

        #gr "cls"
        call draw6BmpLines bgColor$, gridSize * 5
        #gr "getbmp 6NE 0 0 "; str$(gridSize * 5); " "; str$(gridSize * 4);
        bmpsave "6NE", "bmp/6NE.bmp"
    end sub

    sub createHEbmpFiles bgColor$

        #gr "cls"
        call draw1BmpLines bgColor$, gridSize
        #gr "getbmp 1HE 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "1HE", "bmp/1HE.bmp"

        #gr "cls"
        call draw2BmpLines bgColor$, gridSize
        #gr "getbmp 2HE 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "2HE", "bmp/2HE.bmp"

        #gr "cls"
        call draw3BmpLines bgColor$, gridSize
        #gr "getbmp 3HE 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "3HE", "bmp/3HE.bmp"

        #gr "cls"
        call draw4BmpLines bgColor$, gridSize
        #gr "getbmp 4HE 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "4HE", "bmp/4HE.bmp"

        #gr "cls"
        call draw5BmpLines bgColor$, gridSize
        #gr "getbmp 5HE 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "5HE", "bmp/5HE.bmp"

        #gr "cls"
        call draw6BmpLines bgColor$, gridSize * 5
        #gr "getbmp 6HE 0 0 "; str$(gridSize * 5); " "; str$(gridSize * 4);
        bmpsave "6HE", "bmp/6HE.bmp"
    end sub

    sub createNltrBmpFile char$, bgColor$
        #gr "cls"
        loadbmp "2NE", "bmp\2NE.bmp"
        #gr "drawbmp 2NE 0 0"
        #gr "backcolor "; bgColor$
        call drawSmallChar char$, gridSize , gridSize
        #gr "getbmp 2N"; char$; " 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "2N"; char$ , "bmp/2N"; char$; ".bmp"

        #gr "cls"
        loadbmp "3NE", "bmp\3NE.bmp"
        #gr "drawbmp 3NE 0 0"
        #gr "backcolor "; bgColor$
        call drawSmallChar char$, gridSize , gridSize
        #gr "getbmp 3N"; char$; " 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "3N"; char$ , "bmp/3N"; char$; ".bmp"
        
        #gr "cls"
        loadbmp "4NE", "bmp\4NE.bmp"
        #gr "drawbmp 4NE 0 0"
        #gr "backcolor "; bgColor$
        call drawSmallChar char$, gridSize , gridSize
        #gr "getbmp 4N"; char$; " 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "4N"; char$ , "bmp/4N"; char$; ".bmp"

        #gr "cls"
        loadbmp "6NE", "bmp\6NE.bmp"
        #gr "drawbmp 6NE 0 0"
        #gr "backcolor "; bgColor$
        call drawBigChar char$, gridSize * 5, gridSize * 4
        #gr "getbmp 6N"; char$; " 0 0 "; str$(gridSize * 5); " "; str$(gridSize * 4);
        bmpsave "6N"; char$ , "bmp/6N"; char$; ".bmp"
    end sub

    sub createHltrBmpFile char$, bgColor$
        #gr "cls"
        loadbmp "2HE", "bmp\2HE.bmp"
        #gr "drawbmp 2HE 0 0"
        #gr "backcolor "; bgColor$
        call drawSmallChar char$, gridSize , gridSize
        #gr "getbmp 2H"; char$; " 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "2H"; char$ , "bmp/2H"; char$; ".bmp"

        #gr "cls"
        loadbmp "3HE", "bmp\3HE.bmp"
        #gr "drawbmp 3HE 0 0"
        #gr "backcolor "; bgColor$
        call drawSmallChar char$, gridSize , gridSize
        #gr "getbmp 3H"; char$; " 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "3H"; char$ , "bmp/3H"; char$; ".bmp"
        
        #gr "cls"
        loadbmp "4HE", "bmp\4HE.bmp"
        #gr "drawbmp 4HE 0 0"
        #gr "backcolor "; bgColor$
        call drawSmallChar char$, gridSize , gridSize
        #gr "getbmp 4H"; char$; " 0 0 "; str$(gridSize ); " "; str$(gridSize );
        bmpsave "4H"; char$ , "bmp/4H"; char$; ".bmp"

        #gr "cls"
        loadbmp "6HE", "bmp\6HE.bmp"
        #gr "drawbmp 6HE 0 0"
        #gr "backcolor "; bgColor$
        call drawBigChar char$, gridSize * 5, gridSize * 4
        #gr "getbmp 6H"; char$; " 0 0 "; str$(gridSize * 5); " "; str$(gridSize * 4);
        bmpsave "6H"; char$ , "bmp/6H"; char$; ".bmp"
    end sub


    sub drawBigChar char$, width, height
        #gr "font times_new_roman 0 "; str$(int(0.8 * (width)))
        #gr "stringwidth? char$ charWidth"
        #gr "place "; width/2 - charWidth/2; " "; str$(int(0.75 * (height)))
        #gr "\"; char$
    end sub

    sub drawSmallChar char$, width, height
        #gr "font times_new_roman 0 "; str$(int(0.65 * (width)))
        #gr "stringwidth? char$ charWidth"
        #gr "place "; width/2 - charWidth/2; " "; str$(int(0.85 * (height)))
        #gr "\"; char$
    end sub

    sub draw1BmpLines bgColor$, size
        call fillCanvas bgColor$, size
        #gr "size 3"
        #gr "color black"
        ' left side border
        #gr "line "; borderInset; " "; borderInset; " "; borderInset; " "; size
        ' top side bottom
        #gr "line "; borderInset; " "; borderInset; " "; size; " "; borderInset
    end sub

    sub draw2BmpLines bgColor$, size
        call fillCanvas bgColor$, size
        #gr "size 3"
        #gr "color black"
        ' top side border
        #gr "line 0 "; borderInset; " "; size; " "; borderInset
        ' box
        #gr "size 1"
        #gr "place "; 3; " "; boxInset
        #gr "box "; size - 3; " "; size
    end sub

    sub draw3BmpLines bgColor$, size
        call fillCanvas bgColor$, size
        #gr "size 3"
        #gr "color black"
        ' top side border
        #gr "line 0 "; borderInset; " "; size; " "; borderInset
        ' box
        #gr "size 1"
        #gr "place "; 3; " "; boxInset
        #gr "box "; size - 3; " "; size
    end sub

    sub draw4BmpLines bgColor$, size
        call fillCanvas bgColor$, size
        #gr "size 3"
        #gr "color black"
        ' top side border
        #gr "line 0 "; borderInset; " "; size; " "; borderInset
        ' box
        #gr "size 1"
        #gr "place "; 3; " "; boxInset
        #gr "box "; size - 3; " "; size
    end sub

    sub draw5BmpLines bgColor$, size
        call fillCanvas bgColor$, size
        #gr "size 3"
        #gr "color black"
        ' right side border
        #gr "line "; size - borderInset; " "; borderInset; " "; size - borderInset; " "; size
        ' top side border
        #gr "line 0 "; borderInset; " "; size - borderInset; " "; borderInset
    end sub

    sub draw6BmpLines bgColor$, size
        call fillCanvas bgColor$, size
        #gr "size 3"
        #gr "color black"
        height = size * 0.8
        ' left side border
        #gr "line "; borderInset; " 0 "; borderInset; " "; height - borderInset
        ' bottom border
        #gr "line "; borderInset; " "; height - borderInset; " "; size - borderInset; " "; height - borderInset
        ' right side border
        #gr "line "; size - borderInset; " "; height - borderInset; " "; size - borderInset; " 0"
    end sub

    sub fillCanvas bgColor$, size
        #gr "backcolor "; bgColor$
        #gr "color "; bgColor$
        #gr "place 0 0"
        #gr "boxfilled "; size; " "; size
    end sub



