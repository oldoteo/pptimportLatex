'Script utility for windows which exports a PNG or PDF of a given slide from a PowerPoint presentation.
'USAGE FROM COMMAND LINE:
'exportPptToPng <pptfilename IN CURRENT FOLDER> <slidenumber> <0=png OR 1=PDF>
'The script opens the ppt IN THE CURRENT FOLDER and exports the shapes in the given slide IN THE CURRENT Folder.
'The created image is called <pptfilename_including_extension>_<slidenum>.png
'and its modified date coincides with the ppt's modified date.
'To avoid recreating an image for a file which has not been modified, the script first checks whether the modified dates are different
'The exportPptToPng file SHOULD BE ON THE COMMAND PATH of WINDOWS 
'(i.e., opening a command prompt in any folder, typing exportPptToPng (without arguments) and then Enter
'must show a message box with an error, but it must not show "...Unrecognized command name")
'Example: suppose that the folder C:\Users\Io\Desktop contains a powerpoint presentation Figures.pptx with 5 slides
'Open a command prompt (cmd), then navigate to that folder (i.e. cd C:\Users\Io\Desktop)
'Then, the following line should create a PDF file called Figures.pptx_3.pdf containing slide 3: 
'	exportPptToPng Figures.pptx 3 1
'Then, the following line should create a PNG file called Figures.pptx_2.png containing slide 2: 
'	exportPptToPng Figures.pptx 2 0
'
'The generated files will be tightly cropped to the extent of the shapes in the specified slide.
'PDF files contain vector graphics (shapes, text, graphs do not get blurry when zooming in) but due to a Powerpoint restriction, no objects outside the slide extents will be visible.
'PNG files contain bitmapped graphics (colored pixels) and, although overscaled internally, will get blurry when zooming in but will include objects even beyond the slide extents.
'
'This tool can be used with Latex to import a specific slide of a Power Point presentation:
'The exportPptToPng file MUST BE ON THE COMMAND PATH of WINDOWS 
'(i.e., opening a command prompt in any folder, typing exportPptToPng (without arguments) and then Enter
'must show a message box with an error, but it must not show "...Unrecognized command name")
'Declare the following Latex commands in the preable of the Latex document:
'
'\usepackage{graphicx} %Needed somewhere in the preamble, before the \providecommand instructions
'
'\providecommand{\includegraphicspptpdf}[3][]{%
'\immediate\write18{call exportPptToPng.vbs #2 #3 1}%
'\includegraphics[#1]{#2_#3.pdf}
'}
'
'\providecommand{\includegraphicspptpng}[3][]{%
'\immediate\write18{call exportPptToPng.vbs #2 #3 0}%
'\includegraphics[#1]{#2_#3.png}
'}
'
'Then in the document body use, for inserting a vector graphics PDF (which is anyway cropped at maximum to the slide extents, no objects outside the slide extents will be visible):
'\begin{figure}[!hbp]%
'%\centering\includegraphics[width=\columnwidth]{CascadeGeneral.png}%
'\centering\includegraphicspptpdf[width=\columnwidth]{Figures.pptx}{8}%
'\caption{Inline antenna array: a cascade of several 2-port sections closed by a 1-port terminal section}%
'\label{fig:CascadeGeneral}%
'\end{figure}
'
'Or use, in the document body, for inserting a high-res image PNG (which can extend beyond the slide extents and zooming in can become blurry (text and lines too)):
'\begin{figure}[!hbp]%
'%\centering\includegraphics[width=\columnwidth]{CascadeGeneral.png}%
'\centering\includegraphicspptpdf[width=\columnwidth]{Figures.pptx}{8}%
'\caption{Inline antenna array: a cascade of several 2-port sections closed by a 1-port terminal section}%
'\label{fig:CascadeGeneral}%
'\end{figure}
'
'
'
'
'CHANGELOG
'	27/06/2022: updated to support rotated shapes (.Top, .Height, .Width, .Height returned by PowerPoint only refer to the NON-ROTATED SHAPE!)
'	03/2022: created by Matteo Oldoni
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory
'msgbox strCurDir
pptname=WScript.Arguments.Item(0)
slidenum=WScript.Arguments.Item(1)
outputtype_png0_pdf1=0
if WScript.Arguments.Item(2)="1" then
	outputtype_png0_pdf1=1
end if

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
	
	outputname=pptname & "_" & slidenum & "." & extension
	outputname=fso.BuildPath(strCurDir, outputname)
	
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
	if recreate then
		'msgbox(slidenum)
        'Not existing or not uptodate
        Set ppt = CreateObject("PowerPoint.Application")
        ppt.Visible = true
		closepptwhendone=true
		fullname=fso.BuildPath(strCurDir, pptname)
		'msgbox(fso.BuildPath(strCurDir, pptname))
        ppt.Presentations.Open fullname
		if outputtype_png0_pdf1=0 then
			'PNG EXPORT
			'The scaling is done with respect to the slide size (https://stackoverflow.com/questions/36333936/how-to-adjust-export-image-size-in-powerpoint-using-vba)
			'ppt.ActivePresentation.Slides(CInt(slidenum)).Shapes.Range().Export outputname,2, ppt.ActivePresentation.PageSetup.SlideWidth*3, ppt.ActivePresentation.PageSetup.SlideHeight*3, 4
			ppt.ActivePresentation.Slides(CInt(slidenum)).Shapes.Range().Export outputname,extensionid, ppt.ActivePresentation.PageSetup.SlideWidth*3, ppt.ActivePresentation.PageSetup.SlideHeight*3, 4
		end if 
		if outputtype_png0_pdf1=1 then
			''''''EXPORT TO PDF: WORKING BUT THE PDF IS OF THE WHOLE SLIDE: CROPPING STILL TO DO (here or in latex)
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
				if top>newmintop then
					top=newmintop 'In points (1/72 of inch; 1 pixel is 1/96 of an inch)
				end if
				if lft>newminleft then
					lft=newminleft 'In points (1/72 of inch; 1 pixel is 1/96 of an inch)
				end if
				if bottom<newmaxtop then
					bottom=newmaxtop 'In points (1/72 of inch; 1 pixel is 1/96 of an inch)
				end if
				if rght<newmaxleft then
					rght=newmaxleft 'In points (1/72 of inch; 1 pixel is 1/96 of an inch)
				end if
				'msgbox("T=" & shapes.item(i).Top & "; H=" & shapes.item(i).Height & "; L=" & shapes.item(i).Left & "; W=" & shapes.item(i).Width)
			next
			slideheight=ppt.ActivePresentation.PageSetup.SlideHeight 'in pt (1/72 of inch)
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
					Wscript.echo "unable to set date"
				end if 
			end if
		next 
		'ppt.ActivePresentation.Slides(3).Shapes.SelectAll
		'ppt.ActiveWindow.Selection.ShapeRange.Export outputname, 2
        'ppt.ActivePresentation.Slides(1).Export outputname, "PNG"
    end if
end if	

