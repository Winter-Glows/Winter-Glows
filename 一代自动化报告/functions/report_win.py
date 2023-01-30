from pathlib import Path

import pythoncom
from docxtpl import DocxTemplate
from win32com import client


class Word:
    def __init__(self, filepath: str):
        """Parameters
        ----------
        filepath : str
            path of a docx file
        """
        self.filepath = Path(filepath).absolute()

    def to_PDF(self, outfile: str = None, wps: bool = False) -> Path:
        """convert the word to a PDF
        
        Parameters
        ----------
        outfile : str, optional
            path of the generated word file, by default None means the stem name of the word
        wps : bool, optional
            whether use WPS to convert, by default False
        
        Returns
        -------
        Path
            path of the generated word file
        
        Raises
        ------
        OSError
            Only support Windows system 
        """
        

        if outfile:
            outfile = Path(outfile).absolute()
        else:
            outfile = self.filepath.with_suffix(".pdf")
        if not outfile.parent.exists():
            outfile.parent.mkdir(parents=True)
        if wps:
            pythoncom.CoInitialize()
            app = client.Dispatch("Kwps.Application")
            doc = app.Documents.Open(str(self.filepath))
            doc.ExportAsFixedFormat(str(outfile), 17)
            doc.Close()
            app.Quit()
            pythoncom.CoUninitialize()
        else:
            pythoncom.CoInitialize()
            app = client.Dispatch("Word.Application")
            doc = app.Documents.Open(str(self.filepath))
            doc.SaveAs(str(outfile), 17)
            doc.Close()
            app.Quit()
            pythoncom.CoUninitialize()
        return outfile

    def render(self, context: dict, outfile: str) -> Path:
        """use a dict to replace the placeholder in the word 
        
        Parameters
        ----------
        context : dict
            placeholder and value for replacing
        outfile : str
            path of the generated word file
        
        Returns
        -------
        Path
            path of the generated word file
        """
        outfile = Path(outfile).absolute()
        if not outfile.parent.exists():
            outfile.parent.mkdir(parents=True)
        doc = DocxTemplate(str(self.filepath))
        doc.render(context)
        doc.save(str(outfile))
        return outfile

def run(context, CFG, tpl, mode, info, result_dir) -> str:
    # report_model = Word(f"{CFG['Dirs']['templates']}/{mode}/{info['program']}.docx")
    # report = report_model.render(context, f"{result_dir}/{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['zkjy_project']}.docx")
    # Word(report).to_PDF(f"{result_dir}/{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['zkjy_project']}.pdf")
    # return f"{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['zkjy_project']}.pdf"

    report_model = tpl
    report_model.render(context)
    report_model.save(f"{result_dir}/{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['zkjy_project']}.docx")
    docx_path = f"{result_dir}/{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['zkjy_project']}.docx"
    Word(docx_path).to_PDF(f"{result_dir}/{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['zkjy_project']}.pdf")
    return [f"{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['zkjy_project']}.docx", 
            f"{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['zkjy_project']}.pdf"]
    