# pptimportLatex
A Latex package to import pictures directly from Power Point, automating the whole process of exporting, cropping, saving importing in a Latex document.
Works only in Windows
Drawing pictures in Latex is a pain in the ass, even with the available drawing packages, most notably the all-powerful TikZ (a brilliant work: the REAL way to draw pictures in Latex). Mostly, I rather prefer to draw my pictures and diagrams in Microsoft Power Point, and here’s the problem.
Instead of doing it manually, the pptimport package does everything you would do by hand.
Usage is very simple, as it provides 6 commands to import a slide’s content into a Latex figure, providing alternative forms under-the-hood to use a PNG (pixel-based) or a PDF file (for vector content). 
In its simplest form, a figure with
the content of slide 3 of a Power Point presentation figures.pptx is created by

\begin{figure}
\includegraphicspptpdf{figures.pptx}{3}
\end{figure}

Beside the basic import of a slide "as-is", there are commands to also automatically replace a placeholder text in the Power Point slide with specific expressions given in Latex.
This can be done in Power Point (thus retaining all visual properties of the text being replaced but not supporting formulas and other nice Latex typesetting) or in Latex (potentially slightly altering the final look, but leveraging the powerful Latex typesetting and making the figure perfectly blend with the rest of your document).
For instance, if your slide 4 of the Power Point presentation figures.pptx contains a rectangle with a text "TT1" it can be replaced by a formula by:

\begin{figure}
\includegraphicspptpdfsubstex{figures.pptx}{3}{$E=mc^2$}
\end{figure}

The package requires the .sty file and two vbs scripts, which must be accessible.
I recommend putting these files in a folder in the PATH environment variable.

The user guide PDF shows the available commands and some examples.
