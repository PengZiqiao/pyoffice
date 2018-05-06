from pathlib import Path
import win32com.client


def join_ppt(path):
    Application = win32com.client.Dispatch("PowerPoint.Application")
    Application.Visible = True
    prs = Application.Presentations.Add()

    # 遍历所有扩展名为pptx的文件
    files = filter(lambda x: x.match('*.pptx'), Path(path).iterdir())
    for file in files:
        source = Application.Presentations.Open(file)
        dest_slide_index = prs.Slides.Count
        source_slide_start = 1
        source_slide_end = source.Slides.Count
        source.Close()
        # 将源文件的slides复制到目标文件中
        prs.Slides.InsertFromFile(file, dest_slide_index, source_slide_start, source_slide_end)


