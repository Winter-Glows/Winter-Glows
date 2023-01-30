from pathlib import Path
from docxtpl import DocxTemplate
import subprocess


class Word:
    def __init__(self, filepath):
        """Parameters
        ----------
        filepath : str
            path of a docx file
        """
        self.filepath = Path(filepath).absolute()
        self.stem = self.filepath.stem
        self.pdfpath = self.filepath.with_suffix(".pdf")

    def to_PDF(self, outdir: Path) -> Path:
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
        
        if not outdir.exists():
            outdir.mkdir(parents=True)
        res = subprocess.run(["docker", "run", "--rm", "-v", f"/root/streamlit_report/{outdir}:/{outdir}",
                              "docx2pdf:latest", f"/{outdir}", f"/{outdir}/{self.filepath.name}"])
        if not res.returncode:
            return self.pdfpath
        else:
            print(res.stdout)
            raise Exception("docx to pdf failed")

    def render(self, context: dict, outfile: Path) -> Path:
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

def run(context, tpl, mode, info, result_dir) -> str:
    report_model = tpl
    report_model.render(context)
    report_model.save(f"{result_dir}/{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['project']}.docx")
    # docx_path = f"{result_dir}/{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['project']}.docx"
    # Word(docx_path).to_PDF(f"{result_dir}/{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['project']}.pdf")
    return f"{context['zk_id']}_{info['id']}_{context['zk_name']}_{context['project']}.docx"

    
