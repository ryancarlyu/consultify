import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.dml.line import LineFormat
from pptx.enum.dml import MSO_LINE
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from pptx.util import Pt

def marvintable(df,filepath="./Consultify.pptx",cell_width=1.5,cell_height=0.5,\
                slide_title="",cell_font_size=16,title_font_size=30):
  """
  Turns pandas DataFrames into consulting-style PowerPoint slides

  Parameters
  ----------
  df : pandas DataFrame
  
  filepath : str, optional
      The filepath where the resulting PowerPoint slide will be saved.
      Include the PowerPoint filename (default is "./Consultify.pptx").
  
  cell_width : float, optional
      The width of each cell in inches (default is 1.5).
  
  cell_height : float, optional
      The height of each cell in inches (default is 0.5).

  slide_title : str, optional
      The text that serves as the title fo the slide (default is no text).
  
  cell_font_size : int, optional
      The font size of text in each cell (default is 16).
  
  title_font_size : int, optional
      The font size of title text (default is 30).
  """
  prs = Presentation()
  slide = prs.slides.add_slide(prs.slide_layouts[5])

  ncols = df.shape[1]
  nrows = df.shape[0] + 1
  cellwidth = Inches(cell_width)
  cellheight = Inches(cell_height)
  bufferwidth = Inches(0.25)
  left = (prs.slide_width-(ncols * cellwidth + (ncols-1) * bufferwidth))/2
  top = (prs.slide_height-(nrows * cellheight))/2
  slide.shapes.title.text = slide_title
  slide.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
  slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(title_font_size)

  table = slide.shapes.add_table(rows=nrows, cols=ncols * 2 - 1, \
                                left=Inches(left/Inches(1)), \
                                top=Inches(top/Inches(1)), \
                                width=Inches((ncols * cellwidth + (ncols-1) * \
                                              bufferwidth)/Inches(1)), \
                                height=Inches((nrows * cellheight)/Inches(1)))

  for cols in range(ncols * 2 - 1):
    if cols % 2 == 1:
      table.table.columns[cols].width = bufferwidth
    else:
      table.table.columns[cols].width = cellwidth

  for rows in range(nrows):
    for cols in range(ncols * 2 - 1):
      if rows == 0:
        if cols % 2 == 1:
          _set_cell_border(table.table.cell(rows,cols),transparent=True)
          table.table.cell(rows,cols).fill.solid()
          table.table.cell(rows,cols).fill.background()
          continue
        _set_cell_border(table.table.cell(rows,cols))
        table.table.cell(rows,cols).text_frame.add_paragraph
        table.table.cell(rows,cols).fill.solid()
        table.table.cell(rows,cols).fill.background()
        table.table.cell(rows,cols).text_frame.paragraphs[0].text = df.columns[cols//2]
        table.table.cell(rows,cols).text_frame.paragraphs[0].font.color.rgb = RGBColor(0x0,0x0,0x0)
        table.table.cell(rows,cols).text_frame.paragraphs[0].font.size = Pt(cell_font_size)
        table.table.cell(rows,cols).text_frame.word_wrap = True
        table.table.cell(rows,cols).vertical_anchor = MSO_ANCHOR.BOTTOM
      elif rows == nrows - 1:
        _set_cell_border(table.table.cell(rows,cols),transparent=True)
        table.table.cell(rows,cols).text_frame.add_paragraph
        table.table.cell(rows,cols).fill.solid()
        table.table.cell(rows,cols).fill.background()
        if cols % 2 == 1:
          continue
        table.table.cell(rows,cols).text_frame.paragraphs[0].text = str(df.iloc[rows - 1,cols//2])
        makeParaBulletPointed(table.table.cell(rows,cols).text_frame.paragraphs[0])
        table.table.cell(rows,cols).text_frame.paragraphs[0].font.color.rgb = RGBColor(0x0,0x0,0x0)
        table.table.cell(rows,cols).text_frame.paragraphs[0].font.size = Pt(cell_font_size)
        table.table.cell(rows,cols).text_frame.word_wrap = True
        table.table.cell(rows,cols).vertical_anchor = MSO_ANCHOR.MIDDLE
      else:
        _set_cell_border(table.table.cell(rows,cols),border_color='gray',dash='dash')
        table.table.cell(rows,cols).text_frame.add_paragraph
        table.table.cell(rows,cols).fill.solid()
        table.table.cell(rows,cols).fill.background()
        if cols % 2 == 1:
          continue
        table.table.cell(rows,cols).text_frame.paragraphs[0].text = str(df.iloc[rows - 1,cols//2])
        makeParaBulletPointed(table.table.cell(rows,cols).text_frame.paragraphs[0])
        table.table.cell(rows,cols).text_frame.paragraphs[0].font.color.rgb = RGBColor(0x0,0x0,0x0)
        table.table.cell(rows,cols).text_frame.paragraphs[0].font.size = Pt(cell_font_size)
        table.table.cell(rows,cols).text_frame.word_wrap = True
        table.table.cell(rows,cols).vertical_anchor = MSO_ANCHOR.MIDDLE
        
  table.left = Inches((prs.slide_width - table.width)/2/Inches(1))
  table.top = Inches((prs.slide_height - table.height)/2/Inches(1))
  
  if not filepath.endswith(".pptx"):
    filepath = filepath + ".pptx"
  prs.save(filepath)