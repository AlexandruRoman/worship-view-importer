import aspose.slides as slides
import aspose.pydrawing as drawing
import glob

for filename in glob.iglob("/Projects/church/cantari/" + '**/*.ppt', recursive=True):
  try:
    with slides.Presentation(filename) as presentation:
        presentation.save(filename + "x", slides.export.SaveFormat.PPTX)
    print(filename)
  except:
    print("\n\nerror:" + filename)
    