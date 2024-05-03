'Script utility for windows which exports a PNG or PDF of a given slide from a PowerPoint presentation and replaces some placeholder text TT1, TT2...
'See exportPptToPng.vbs for the general principle.
'This script uses the syntax
'exportPptToPngReplace <pptfilename IN CURRENT FOLDER> <slidenumber> <0=png OR 1=PDF> <outputnameappend> <replacementlist> <0=inPowerPoint OR 1=inLatex>
'where replacementlist is replacestr1;replacestr2;... (semicolon-separated list of replacement strings)
'Each string is replaced to TT1, TT2, ... in the Power Point file
'
'
'CHANGELOG
'	02/05/2024: updated to support replacement text which will be replaced in the given slide to "t1", "t2"...
'	02/05/2024: updated to support an additional filename append 
'	27/06/2022: updated to support rotated shapes (.Top, .Height, .Width, .Height returned by PowerPoint only refer to the NON-ROTATED SHAPE!)
'	03/2022: created by Matteo Oldoni
'msgbox "N"
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory
'msgbox strCurDir
pptname=WScript.Arguments.Item(0)
slidenum=WScript.Arguments.Item(1)
outputtype_png0_pdf1=0
if WScript.Arguments.Item(2)="1" then
	outputtype_png0_pdf1=1
end if
outputfilenameappend=""
if WScript.Arguments.Count>3 Then
	outputfilenameappend=WScript.Arguments.Item(3)
end if
'Replacement strings
replacedToBePrefix="TT"
dim replacementTextArray
if WScript.Arguments.Count>4  Then
	'numreplacements=WScript.Arguments.Count-4
	'redim replacementTextArray(numreplacements-1)
	'for ii=0 to numreplacements-1
'		replacementTextArray(ii)=WScript.Arguments.Item(ii+4)
	'next
	replacementTextArray = Split (WScript.Arguments.Item(4),";")  
	'msgbox "UBound=" & UBound(replacementTextArray)
else
	replacementTextArray=Split("",";") 'Just to create an empty array
	'msgbox "Empty" & WScript.Arguments.Count
end if 

replaceInTex=False
if WScript.Arguments.Count>5  Then
 if CInt(WScript.Arguments.Item(5))>0 then
	replaceInTex=True
 end if
end If

set fso = CreateObject("Scripting.FileSystemObject")
'  ppShapeFormatGIF = 0  '&H0
'  ppShapeFormatJPG = 1  '&H1
'  ppShapeFormatPNG = 2  '&H2
'  ppShapeFormatBMP = 3  '&H3
'  ppShapeFormatWMF = 4  '&H4
'  ppShapeFormatEMF = 5  '&H5
extension="png"
extensionid=2
if outputtype_png0_pdf1=1 then
	extension="pdf"
	extensionid=0
end if 
if fso.FileExists(pptname) then
	Set objFile = fso.GetFile(pptname)
    datetim = objFile.DateLastModified
	
	outputname=pptname & "_" & slidenum & outputfilenameappend & "." & extension
	'msgbox outputname
	outputname=fso.BuildPath(strCurDir, outputname)
	
	texfilename=outputname & ".t" 'Used only when replaceInTex=True
	'WScript.echo imagename
	recreate=true
    if fso.FileExists(outputname)	then
        'Already existing: check if it is updated by comparing the modifydate
		Set objFile = fso.GetFile(outputname)
		datetimimage = objFile.DateLastModified
        if datetimimage<>datetim then
            'Image is not corresponding: update it
			'WScript.StdOut.Write "Already exists but not uptodate: " & outputname & " has date " & datetimimage & " while ppt is " & datetim
			recreate=true
		else
			recreate=false
		end if	
    end if
	if recreate=False and replaceInTex then
		'the pdf/png is already there.
		'But check if the .t file is also there
		if fso.FileExists(texfilename)	then
			'Already existing: check if it is updated by comparing the modifydate
			Set objFile = fso.GetFile(texfilename)
			datetimimage = objFile.DateLastModified
			if datetimimage<>datetim then
				'Image is not corresponding: update it
				'WScript.StdOut.Write "Already exists but not uptodate: " & texfilename & " has date " & datetimimage & " while ppt is " & datetim
				recreate=true
			else
				recreate=false
			end if	
		end if		
	end if 
	
	if recreate then
		'msgbox(slidenum)
        'Not existing or not uptodate
        Set ppt = CreateObject("PowerPoint.Application")
        ppt.Visible = true
		closepptwhendone=true
		fullname=fso.BuildPath(strCurDir, pptname)
		'msgbox(fso.BuildPath(strCurDir, pptname))
        'ppt.Presentations.Open fullname, False 'Write/read
		ppt.Presentations.Open fullname, True 'Readonly; Tried setting WithWindow=False but there is a subsequent error for "There is no active presentation"
		'expression.Open (FileName, ReadOnly, Untitled, WithWindow)
		'Replace text
		dim replacedCoordinatesShapeId()
		redim replacedCoordinatesShapeId(UBound(replacementTextArray,1))
		for ir=0 to UBound(replacementTextArray,1)
			Set oSld = ppt.ActivePresentation.Slides(CInt(slidenum))
			For isp=0 to oSld.Shapes.Count-1 'Each oShp In oSld.Shapes 
				set oShp=oSld.Shapes.item(isp+1)  'Load the current shape
				'msgbox("oShp")
				if not oShp is Nothing Then
					if oShp.HasTextFrame then
					if not oShp.TextFrame is Nothing then
						if oShp.TextFrame.HasText then
						if not oShp.TextFrame.TextRange is Nothing then
							Set oTxtRng = oShp.TextFrame.TextRange
							'msgbox("ir=" & ir & "     UBound=" & UBound(replacementTextArray,1) & "    replacing=" & replacedToBePrefix & CInt(ir+1))
							if replaceInTex Then
								'Just remove text placeholder in powerpoint and store its shapeid
								Set oTmpRng = oTxtRng.Replace(replacedToBePrefix & CInt(ir+1), "   ",0,True,True) 'Replace with 3 spaces so that it occupies the same space as "TTn"
								if not oTmpRng Is nothing then
									'Some text has been replaced: store the shape id, so that we will later retrieve its coordinates and convert them into relative coordinates and append them to a .tex file containing the \put commands
									replacedCoordinatesShapeId(ir)=isp+1 'This is the actual 1-based index of the shape
								end if 								
							else
								'Replace text directly into power point
								Set oTmpRng = oTxtRng.Replace(replacedToBePrefix & CInt(ir+1), replacementTextArray(ir),0,True,True)
								'Continue replacing
								Do While Not oTmpRng Is Nothing
									Set oTxtRng = oTxtRng.Characters(oTmpRng.Start + oTmpRng.Length, oTxtRng.Length)
									Set oTmpRng = oTxtRng.Replace(replacedToBePrefix & CInt(ir+1), replacementTextArray(ir),0,True,True)
								Loop
							end if 
						end if
						end if
					end if
					end if
				end if 
			Next
		next

		
		'Now we need to write the replacements (needed when using replaceinTex) and the borders computation (borders are needed to crop the pdf which otherwise is the whole slide)
		set shapes=ppt.ActivePresentation.Slides(CInt(slidenum)).Shapes.Range()
		'msgbox("T=" & shapes.Top & "; H=" & shapes.Height & "; L=" & shapes.Left & "; W=" & shapes.Width)
		''Get size of the current shapes
		top=100000
		lft=100000
		bottom=-100000
		rght=-100000
		for i = 1 to shapes.Count
			'Compute effect of rotation (Z, in clockwise sense)
			'TO do this, we compute the rotated coordinates of the four corners (At: point A top coordinate; Al: point A left coordinate; and so on)
			set shp=shapes.item(i)
			rot_rad=-shp.Rotation*3.14159265/180 'shp.Rotation is in clockwise degrees. rot_rad instead is converted to radians in counterclockwise direction
			'Here we retrieve the coordinates of the box vertices accounting for rotation, but the Latex command NEGLECT ROTATION AND ALWAYS USES THE COMPUTED COORDINATES OF A and B VERTICES of the shape
			At=shp.Top+shp.Height/2 +(-shp.Height/2)*cos(rot_rad) -(-shp.Width/2)*sin(rot_rad)
			Al=shp.Left+shp.Width/2 +(-shp.Height/2)*sin(rot_rad) +(-shp.Width/2)*cos(rot_rad)
			Bt=shp.Top+shp.Height/2 +(-shp.Height/2)*cos(rot_rad) -( shp.Width/2)*sin(rot_rad)
			Bl=shp.Left+shp.Width/2 +(-shp.Height/2)*sin(rot_rad) +( shp.Width/2)*cos(rot_rad)
			Ct=shp.Top+shp.Height/2 +( shp.Height/2)*cos(rot_rad) -( shp.Width/2)*sin(rot_rad)
			Cl=shp.Left+shp.Width/2 +( shp.Height/2)*sin(rot_rad) +( shp.Width/2)*cos(rot_rad)
			Dt=shp.Top+shp.Height/2 +( shp.Height/2)*cos(rot_rad) -(-shp.Width/2)*sin(rot_rad)
			Dl=shp.Left+shp.Width/2 +( shp.Height/2)*sin(rot_rad) +(-shp.Width/2)*cos(rot_rad)				
			'Now select the rectangle which holds the extreme coordinates of the rotated corner points
			'newmintop must be set as the smallest distance from the top border to this object
			newmintop=At
			if Bt<newmintop then
				newmintop=Bt
			end if
			if Ct<newmintop then
				newmintop=Ct
			end if
			if Dt<newmintop then
				newmintop=Dt
			end if
			'Compute also any border. Sometimes some shapes return width=-1e9 (points) which then creates overflow, so we add a check
			lineweight=0
			'msgbox("i=" & i & " Newmintop = " & newmintop)
			'set lin=shapes.item(i).Line
			'msgbox("i=" & i & " Newmintop = " & newmintop & "lineok")
			weight=0 'Some objects moreover DO NOT HAVE the .Line.Weight property
			on error resume next
			weight=shapes.item(i).Line.Weight
			on error goto 0
			'msgbox("i=" & i & " Newmintop = " & newmintop & "; width=" & weight)
			if weight<1e4 then
				if weight>-1e4 then
					lineweight=weight
				end if
			end if 
			'msgbox("i=" & i & " Newmintop = " & newmintop & "; width=" & shapes.item(i).Line.Weight)
			newmintop=newmintop-(lineweight/2)
			'msgbox("Newmintop updated = " & newmintop & "; width=" & shapes.item(i).Line.Weight)	
			'newmaxtop must be set as the largest distance from the top border to this object
			newmaxtop=At  
			if Bt>newmaxtop then
				newmaxtop=Bt
			end if
			if Ct>newmaxtop then
				newmaxtop=Ct
			end if
			if Dt>newmaxtop then
				newmaxtop=Dt
			end if		
			newmaxtop=newmaxtop+lineweight/2

			newmaxleft=Al
			if Bl>newmaxleft then
				newmaxleft=Bl
			end if
			if Cl>newmaxleft then
				newmaxleft=Cl
			end if
			if Dl>newmaxleft then
				newmaxleft=Dl
			end if		
			newmaxleft=newmaxleft+lineweight/2

			newminleft=Al
			if Bl<newminleft then
				newminleft=Bl
			end if
			if Cl<newminleft then
				newminleft=Cl
			end if
			if Dl<newminleft then
				newminleft=Dl
			end if					
			newminleft=newminleft-lineweight/2
			
			'Now we know the horizontal bounding box of this shape: update the extent of the overall bounding box
			'top is the size of the top border (from the top side of the slide to the topmost vertex of the topmost object
			if top>newmintop then
				top=newmintop 'In points (1/72 of inch; 1 pixel is 1/96 of an inch)
			end if
			'lft is the size of the left border (from left edge of the slide to the leftmost vertex of the leftmost object)
			if lft>newminleft then
				lft=newminleft 'In points (1/72 of inch; 1 pixel is 1/96 of an inch)
			end if
			'bottom is the distance from the top edge of the slide to the bottomest vertex of the bottomest object
			if bottom<newmaxtop then
				bottom=newmaxtop 'In points (1/72 of inch; 1 pixel is 1/96 of an inch)
			end if
			'rght is the distance from the left edge of the slide to the rightmost vertex of the rightmost object
			if rght<newmaxleft then
				rght=newmaxleft 'In points (1/72 of inch; 1 pixel is 1/96 of an inch)
			end if
			'msgbox("T=" & shapes.item(i).Top & "; H=" & shapes.item(i).Height & "; L=" & shapes.item(i).Left & "; W=" & shapes.item(i).Width)
		next
		slideheight=ppt.ActivePresentation.PageSetup.SlideHeight 'in pt (1/72 of inch)
		slidewidth=ppt.ActivePresentation.PageSetup.SlideWidth 'in pt (1/72 of inch)
		'Here we have the size of the slide in pt (slideheight, slidewidth)
		'We also have the various borders: top, lft, bottom, rght
		if replaceInTex Then
			'We are doing a replacement of strings in Tex: hence we already deleted the TTn text from the PPT boxes but we still have to generate a tex file with the coordinates to write the Latex replacements
			'The coordinates of the replaced shapes must be translated to permil and appended in a .tex file with \put(x,y){text}
			'Beware: pdf clips stuff outside of the slide, whereas png retains everything
			if outputtype_png0_pdf1=0 then
				contentheight=bottom-top
				contentwidth=rght-lft
				milpermil=CDbl(contentheight) 'Height is the 1000 permil
				if contentwidth>contentheight then
					milpermil=CDbl(contentwidth) 'Width is the 1000 permil
				end if
			else			
				'exporting pdf: do not consider anything beyond the slide borders
				if bottom>SlideHeight then
					'Some content was beyond the lower edge of the slide: it will be clipped by the pdf export
					bottom=slideheight
				end if 
				if bottom<0 then
					bottom=0 'All content was above the slide top edge: it will be clipped by the pdf EXPORT
				end if 
				if top<0 then
					'Some content was above the top edge of the slide: it will be clipped by the pdf export
					top=0
				end if 
				if top>SlideHeight then
					top=SlideHeight 'All content was below the slide bottom edge: it will be clipped by the pdf EXPORT
				end if 					
				if lft<0 then
					'Some content was outside to the left: it will be clipped by pdf EXPORT
					lft=0
				end if
				if lft>SlideWidth then
					'All content was outside to the right: it will be clipped by pdf EXPORT
				end if 
				if rght>slidewidth then
					'Some content was outside to the right: it will be clipped by the pdf EXPORT
					rght=SlideWidth
				end if
				if rght<0 then
					'all content was outside to the left: it will be clipped by the pdf EXPORT
					rght=0
				end if
				contentheight=bottom-top
				contentwidth=rght-lft
				milpermil=CDbl(contentheight) 'Height is the 1000 permil
				if contentwidth>contentheight then
					milpermil=CDbl(contentwidth) 'Width is the 1000 permil
				end if
				
			end if
			if fso.FileExists(texfilename) then
				fso.DeleteFile  texfilename, True 
			end if
			for irep=0 to UBound(replacedCoordinatesShapeId)  'Loop for each replaced shape ic
				'Get the size, rotation... of this SHAPE
				if replacedCoordinatesShapeId(irep)>0  then
					'Ok, there is a shape in which text has been replaced for this TTn 
					set shp=shapes.item(replacedCoordinatesShapeId(irep))
					rot_rad=-shp.Rotation*3.14159265/180 'shp.Rotation is in clockwise degrees. rot_rad instead is converted to radians in counterclockwise direction
					'Here we retrieve the coordinates of the box vertices accounting for rotation, but the Latex command NEGLECT ROTATION AND ALWAYS USES THE COMPUTED COORDINATES OF A and B VERTICES of the shape
					At=shp.Top+shp.Height/2 +(-shp.Height/2)*cos(rot_rad) -(-shp.Width/2)*sin(rot_rad)
					Al=shp.Left+shp.Width/2 +(-shp.Height/2)*sin(rot_rad) +(-shp.Width/2)*cos(rot_rad)
					Bt=shp.Top+shp.Height/2 +(-shp.Height/2)*cos(rot_rad) -( shp.Width/2)*sin(rot_rad)
					Bl=shp.Left+shp.Width/2 +(-shp.Height/2)*sin(rot_rad) +( shp.Width/2)*cos(rot_rad)
					Ct=shp.Top+shp.Height/2 +( shp.Height/2)*cos(rot_rad) -( shp.Width/2)*sin(rot_rad)
					Cl=shp.Left+shp.Width/2 +( shp.Height/2)*sin(rot_rad) +( shp.Width/2)*cos(rot_rad)
					Dt=shp.Top+shp.Height/2 +( shp.Height/2)*cos(rot_rad) -(-shp.Width/2)*sin(rot_rad)
					Dl=shp.Left+shp.Width/2 +( shp.Height/2)*sin(rot_rad) +(-shp.Width/2)*cos(rot_rad)
					'msgbox "L=" & shp.Left & ", T=" & shp.Top & ", W=" & shp.Width & ", H=" & shp.Height & ", At=
					'Determine the desired alignment
					alignment_minus1right_0center_1left=2 'Unassigned initially
					texttowrite=""
					if Instr(1,replacementTextArray(irep),"R_")=1 then
						'but text is right: right
						alignment_minus1right_0center_1left=-1
						texttowrite=Mid(replacementTextArray(irep),3) 'Mid(string,firstcharactertoextract1base,optionalNumOfChars
					elseif Instr(1,replacementTextArray(irep),"C_")=1 then
						'but text is centered: centered
						alignment_minus1right_0center_1left=-1
						texttowrite=Mid(replacementTextArray(irep),3) 'Mid(string,firstcharactertoextract1base,optionalNumOfChars
					elseif Instr(1,replacementTextArray(irep),"L_")=1 then
						'but text is left: left
						alignment_minus1right_0center_1left=-1
						texttowrite=Mid(replacementTextArray(irep),3) 'Mid(string,firstcharactertoextract1base,optionalNumOfChars
					else
						'Text does not specify alignment: retrieve from shape properties
						texttowrite=replacementTextArray(irep)
					end if
					texttowrite=Replace(texttowrite, "^^", "^") 'Important: remove any double superscript symbol, because these are introduced by the Windows CALL function as automatic escape
					'Now retrieve the borders of the shape, and (if not specified by the text)	
					marginx=0
					marginytop=0
					If shapes.item(replacedCoordinatesShapeId(irep)).HasTextFrame Then
						'Detect the alignment of the text box: Right, Left, Center, based on the alignment of the first paragraph of the text
						'ppAlignCenter	2	Center align;
						'ppAlignDistribute	5	Distribute
						'ppAlignJustify	4	Justify
						'ppAlignJustifyLow	7	Low justify
						'ppAlignLeft	1	Left aligned
						'ppAlignmentMixed	-2	Mixed alignment
						'ppAlignRight	3	Right-aligned
						'ppAlignThaiDistribute 6   Thai-distributed
						if shapes.item(replacedCoordinatesShapeId(irep)).TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Alignment=2 then
							'box is centered
							if alignment_minus1right_0center_1left>1 then
								alignment_minus1right_0center_1left=0		'If unassigned, set to center
							end if
							marginx=0
							marginytop=shapes.item(replacedCoordinatesShapeId(irep)).TextFrame.MarginTop
						elseif shapes.item(replacedCoordinatesShapeId(irep)).TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Alignment=3 then
							'Right
							if alignment_minus1right_0center_1left>1 then
								alignment_minus1right_0center_1left=-1   'If unassigned, set to right
							end if
							marginx=shapes.item(replacedCoordinatesShapeId(irep)).TextFrame.MarginRight
							marginytop=shapes.item(replacedCoordinatesShapeId(irep)).TextFrame.MarginTop
						else
							'Everything else: assume left alignment
							if alignment_minus1right_0center_1left>1 then
								alignment_minus1right_0center_1left=1	'If unassigned, set to left
							end if
							marginx=shapes.item(replacedCoordinatesShapeId(irep)).TextFrame.MarginLeft
							marginytop=shapes.item(replacedCoordinatesShapeId(irep)).TextFrame.MarginTop
						end if
					end if
					boxtext=""
					if alignment_minus1right_0center_1left=-1 then
						'Right align: use the B point of the shape
						xpermilleft=(Bl-marginx*cos(rot_rad)+marginytop*sin(rot_rad)-lft)/milpermil*1000
						ypermilbottom=(bottom-Bt-marginytop*cos(rot_rad)-marginx*sin(rot_rad))/milpermil*1000												
						'boxtext="\raisebox{-\height}[0pt][0pt]{\makebox[0pt][r]{" & texttowrite & "}}" 
						boxtext="\rotatebox[origin=lT]{" & CInt(rot_rad*180/3.14159265) &"}{" & "\raisebox{-\height}[0pt][0pt]{\makebox[0pt][r]{" & texttowrite & "}}" & "}"						
						'msgbox "Right aligning at " & xpermilleft & ", " & ypermilbottom & ", Bl=" & Bl/milpermil*1000 & ", marginx=" & marginx/milpermil*1000 & ", lft=" & lft/milpermil*1000							
					elseif alignment_minus1right_0center_1left=0 then
						'Center align: use the center between A and B points of the shape
						xpermilleft=(Bl/2+Al/2-lft)/milpermil*1000
						ypermilbottom=(bottom-Bt/2-At/2-marginytop*cos(rot_rad)+marginx*sin(rot_rad))/milpermil*1000												
						'boxtext="\raisebox{-\height}[0pt][0pt]{\makebox[0pt][c]{" & texttowrite & "}}"			
						boxtext="\rotatebox[origin=lT]{" & CInt(rot_rad*180/3.14159265) &"}{" & "\raisebox{-\height}[0pt][0pt]{\makebox[0pt][c]{" & texttowrite & "}}" & "}"						
						'msgbox "Center aligning at " & xpermilleft & ", " & ypermilbottom & ", AlBl=" & (Al/2+Bl/2)/milpermil*1000 & ", lft=" & lft/milpermil*1000							
					else
						'Left align: use the A point of the shape and rotate the margins accordingly
						xpermilleft=(Al+marginx*cos(rot_rad)+marginytop*sin(rot_rad)-lft)/milpermil*1000
						ypermilbottom=(bottom-At-marginytop*cos(rot_rad)+marginx*sin(rot_rad))/milpermil*1000												
						'boxtext="\raisebox{-\height}[0pt][0pt]{\makebox[0pt][l]{" & texttowrite & "}}"
						'rot_rad=0
						boxtext="\rotatebox[origin=lT]{" & CInt(rot_rad*180/3.14159265) &"}{" & "\raisebox{-\height}[0pt][0pt]{\makebox[0pt][l]{" & texttowrite & "}}" & "}"
						'boxtext="\fbox{" & texttowrite & "}"	
						'msgbox "Left aligning at " & xpermilleft & ", " & ypermilbottom & ", Al=" & Al/milpermil*1000 & ", marginx=" & marginx/milpermil*1000 & ", lft=" & lft/milpermil*1000 & ", shLft=" & shapes.item(replacedCoordinatesShapeId(irep)).Left/milpermil*1000							
						'msgbox texttowrite
					end if
					'msgbox "W=" & slidewidth & ", H=" & slideheight & ", 1000x1000=" & milpermil & ", At=" & At & ", bottom=" & bottom & ", ypermilbottom=" & ypermilbottom 
					'Set fs = CreateObject("Scripting.FileSystemObject")
					ForAppending=8
					Set f = fso.OpenTextFile(texfilename, ForAppending, True, False)
					'msgbox boxtext
					f.WriteLine "\put(" & CInt(xpermilleft) & "," & CInt(ypermilbottom) & ")" & "{" & boxtext & "}" & "%"
					f.Close
				end if
			next
		end if
		if outputtype_png0_pdf1=0 then
			''''EXPORTING IN PNG
			'The scaling is done with respect to the slide size (https://stackoverflow.com/questions/36333936/how-to-adjust-export-image-size-in-powerpoint-using-vba)
			'ppt.ActivePresentation.Slides(CInt(slidenum)).Shapes.Range().Export outputname,2, ppt.ActivePresentation.PageSetup.SlideWidth*3, ppt.ActivePresentation.PageSetup.SlideHeight*3, 4
			ppt.ActivePresentation.Slides(CInt(slidenum)).Shapes.Range().Export outputname,extensionid, ppt.ActivePresentation.PageSetup.SlideWidth*3, ppt.ActivePresentation.PageSetup.SlideHeight*3, 4
		end if 		
		if outputtype_png0_pdf1=1 then
			'''''EXPORTING PDF: CROP the area to the actual content computed before
			set pres = ppt.ActivePresentation
			set printOptions = pres.PrintOptions
			set range = printOptions.Ranges.Add(slidenum,slidenum) 
			printOptions.RangeType = 4		
			pres.ExportAsFixedFormat outputname, 2, 2, 0, 2, 1, 0, range, 4, "", False, True, False, False, False
			'msgbox("Exported")
			''Now crop the pdf page to the shapes
			const adTypeText=2
			const adTypeBinary=1			
			set inStream=WScript.CreateObject("ADODB.Stream")
			inStream.Open
			inStream.type=adTypeBinary
			inStream.LoadFromFile(outputname)
			dim buff1
			buff1 = inStream.Read()
			''COnvert /MediaBox to byte array to look for:
			set tempStream=WScript.CreateObject("ADODB.Stream")
			tempStream.Open
			tempStream.type=adTypeText
			tempStream.Charset="ASCII"
			tempStream.WriteText("/MediaBox")
			tempStream.Position=0
			tempStream.type=adTypeBinary
			cpboxtext=tempStream.Read()
			tempStream.Close
			'msgbox("cpboxtextLen=" & UBound(cpboxtext))  Returns 9 bytes (LBound=0, UBound=8)
			idofmediabox=InstrB(buff1, cpboxtext) 'DOES NOT FIND ANYTHING when reading ReadText. It works when using the ASCII codes read via binary
			if idofmediabox=0 then
				'UNABLE TO CROP: do nothing
				inStream.Close()
				msgbox("Error: unable to crop the PDF file to shape size")
			end if
			if idofmediabox<>0 then
				'Mediabox found: insert a cropbox
				'msgbox(idofmediabox)
				'Copy first part of pdf file
				set outStream=WScript.CreateObject("ADODB.Stream")
				outStream.Open
				outStream.type=adTypeBinary
				inStream.Position=0
				'for ii=0 to idofmediabox-1
					'msgbox("id=" & idofmediabox & " ii=" & ii & " T=" & typename(buff1) & " L=" & LBound(buff1) & " U=" & UBound(buff1))    'Type is Byte() but not compatible with indexing
					'dim a
					'a=Mid(buff1, ii, 1)  'Required to use 
					'msgbox(a)
					'outStream.Write(chrB((a))) 'Type mismatch
					'outStream.WriteText(((a)))
					'outStream.WriteText(a)
					
				'next 
				outStream.Write inStream.Read(idofmediabox-1)
				'outStream.SaveToFile pdfname & "rew.pdf",2
				'outStream.Close()
				
				'Create cropbox and insert it
				set tempStream=WScript.CreateObject("ADODB.Stream")
				tempStream.Open
				tempStream.Charset="ASCII"
				tempStream.type=adTypeText    '/CropBox[A B C D]: increasing B cuts away part of the bottom; reducing D cuts away part of the top
				'tempStream.WriteText("/CropBox[" & CInt(lft/72*96) & " " & CInt((slideheight-bottom)/72*96) & " " & CInt(rght/72*96)  & " " & CInt((slideheight-top)/72*96) & "] ")
				tempStream.WriteText("/CropBox[" & CInt(lft) & " " & CInt((slideheight-bottom)) & " " & CInt(rght)  & " " & CInt((slideheight-top)) & "] ")
				tempStream.Position=0
				tempStream.type=adTypeBinary
				dim cpboxtext
				cpboxtext=tempStream.Read()
				outStream.Write(cpboxtext)
				tempStream.Close
				outStream.Write inStream.Read()
				inStream.Close()
				outStream.SaveToFile outputname,2
				outStream.Close()
			end if
			'msgbox "stop"
		end if
		''''''END OF PDF EXPORT
		
		ppt.ActivePresentation.Close
		if ppt.Presentations.Count=0 then
		   ppt.Quit
		end if 
		Set objFile = fso.GetFile(outputname)
		'Set modified date to ppt's date
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.NameSpace(objFile.ParentFolder.Path)
		for ii=1 to 3
			objFolder.Items.Item(objFile.Name).ModifyDate = datetim
			WScript.Sleep 50
			Set objFile = fso.GetFile(outputname)
			datetimimage = objFile.DateLastModified
			if datetimimage=datetim then
				ii=3
				'Wscript.echo "Set date for " & outputname
				'Wscript.echo datetimimage
				exit for
			else
				WScript.Sleep 100
				if ii=3 then
					Wscript.echo "unable to set date of " & outputname
				end if 
			end if
		next 
		if replaceInTex then
			'Update timestamp also of the auxiliary texfile (containing the \put commands)
			Set objFile = fso.GetFile(texfilename)
			for ii=1 to 3
			objFolder.Items.Item(objFile.Name).ModifyDate = datetim
			WScript.Sleep 50
			Set objFile = fso.GetFile(texfilename)
			datetimimage = objFile.DateLastModified
			if datetimimage=datetim then
				ii=3
				'Wscript.echo "Set date for " & texfilename
				'Wscript.echo datetimimage
				exit for
			else
				WScript.Sleep 100
				if ii=3 then
					Wscript.echo "unable to set date of " & texfilename
				end if 
			end if
		next 
		end if
		'ppt.ActivePresentation.Slides(3).Shapes.SelectAll
		'ppt.ActiveWindow.Selection.ShapeRange.Export outputname, 2
        'ppt.ActivePresentation.Slides(1).Export outputname, "PNG"
    end if
end if	

