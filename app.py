import io
import os
import platform
import ssl
import zipfile

import camelot as cam
import fitz
import pikepdf
import pytesseract
import wget
from pdf2image import convert_from_path
from PIL import Image


class PDFEXTRACTION:
    """
    A Class is created to extract the following things from a pdf
    Text,Image,Table

    """

    def __init__(self, filename):
        """
        A method is created for initialize the filename and the password

        """
        self.filename = filename
        # self.pass_word = pass_word

    def decrypt(self, password):
        """
        A method is created for decrypting a pdf

        """
        self.password = password
        with pikepdf.open(self.filename, password=self.password) as pdf:
            pdf.save('output.pdf')

    def extracting_images(self, pagenumber=None):
        """
        A method is created for extracting all the images present in the pdf
        Pass Pageno of which you want to extract
        For extracting from all the pages leave it blank

        """

        self.pagenum = pagenumber

        image_dir = './image'

        # open the file
        pdf_file = fitz.open("output.pdf")
        len_pdf = len(pdf_file)
        # # iterate over PDF pages
        # for page in self.pagenumber:
        # get the page itself

        if pagenumber is None:

            for page_index in range(len_pdf):
                # get the page itself

                page = pdf_file[page_index]
                image_list = page.getImageList()
                # printing number of images found in this page
                if image_list:
                    print(
                        f"[+]Found{len(image_list)}images in page{page_index}"
                        )
                else:
                    print("[!] No images found on page", page_index)

                for image_ind, img in enumerate(page.getImageList(), start=1):
                    # get the XREF of the image
                    xref = img[0]
                    # extract the image bytes
                    base_image = pdf_file.extract_image(xref)
                    image_bytes = base_image["image"]
                    # get the image extension
                    image_ext = base_image["ext"]
                    # load it to PIL
                    image = Image.open(io.BytesIO(image_bytes))

                    # save to local disk
                    file_name = f"image{page_index+1}_{image_ind}.{image_ext}"
                    filePath = os.path.join(image_dir, file_name)

                    image.save(open(filePath, "wb"))

        else:

            page = pdf_file[self.pagenum]

            image_list = page.getImageList()
            # printing number of images found in this page
            if image_list:
                print(
                    f"[+]Found {len(image_list)} images in page {self.pagenum}"
                    )
            else:
                print("[!] No images found on page", self.pagenum)
            for image_index, img in enumerate(page.getImageList(), start=1):
                # get the XREF of the image
                xref = img[0]
                # extract the image bytes
                base_image = pdf_file.extract_image(xref)
                image_bytes = base_image["image"]
                # get the image extension
                image_ext = base_image["ext"]
                # load it to PIL
                image = Image.open(io.BytesIO(image_bytes))

                # save it to local disk
                file_name = f"image{self.pagenum+1}_{image_index}.{image_ext}"
                filePath = os.path.join(image_dir, file_name)

                image.save(open(filePath, "wb"))

    def check_directory(self):
        """
        Check whether the specified path is an existing filpath
        And checking the type of operating sestem

        """
        osname = platform.system()

        ssl._create_default_https_context = ssl._create_unverified_context

        if osname == "Windows":
            # Path of the folder
            path = './Tesseract-OCR'

            isExist = os.path.exists(path)

            if not isExist:

                url = "https://osdn.net/frs/g_redir.php?m=nchc&f=tesseract-ocr-alt%2Ftesseract-ocr-3.02-win32-portable.zip"

                wget.download(url, out="tesseract.zip")

                with zipfile.ZipFile("tesseract.zip", 'r') as zip_ref:
                    zip_ref.extractall("./")

                os.remove("tesseract.zip")

            else:
                pass

            path2 = "./poppler-21.11.0"
            isExist2 = os.path.exists(path2)

            if not isExist2:
                url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v21.11.0-0/Release-21.11.0-0.zip"

                wget.download(url, out="poppler.zip")

                with zipfile.ZipFile("poppler.zip", 'r') as zip_ref:
                    zip_ref.extractall("./")

                os.remove("poppler.zip")

            else:
                pass

        elif osname == "Linux":
            pass

    def extracting_text(self, pageno=None):

        """
        The method is extracting all the text from the pdf
        And saved in a text file
        Pass Pageno of which you want to extract
        For extracting from all the pages leave it blank
        """

        pop_path = ".\\poppler-21.11.0\\Library\\bin"

        if platform.system() == "Linux":

            pop_path = "/usr/bin"

        if platform.system() == "Windows":
            path_tes = ".\\Tesseract-OCR\\tesseract.exe"
            pytesseract.pytesseract.tesseract_cmd = r"{0}".format(path_tes)

        pages = convert_from_path(
                "output.pdf", 500,
                poppler_path=r"{0}".format(pop_path)
        )

        image_counter = 1

        for page in pages:
            # Saving pages as image
            filename = "page_"+str(image_counter)+".jpg"
            page.save(filename, 'JPEG')

            image_counter = image_counter + 1

        filelimit = image_counter-1

        outfile = "out_text.txt"

        f = open(outfile, "a", encoding="utf-8")

        if pageno is None:

            for i in range(1, filelimit + 1):

                filename = "page_"+str(i)+".jpg"

                # Taking Text from images
                text = str(((
                    pytesseract.image_to_string(Image.open(filename))
                    )))

                text = text.replace('-\n', '')

                f.write(text)

            f.close()

        else:
            self.pageno = pageno

            filename = "page_"+str(self.pageno)+".jpg"

            # Taking Text from images
            text = str(((pytesseract.image_to_string(Image.open(filename)))))

            text = text.replace('-\n', '')

            f.write(text)

            f.close()

        cwd = os.getcwd()

        li = os.listdir(cwd)

        for i in li:
            if i.endswith(".jpg"):
                os.remove(i)

    def table_pdf(self):
        """
        The method is extracting all the tables present in the PDF
        And saving in a CSV file

        """

        table = cam.read_pdf("output.pdf", pages='1', flavor='stream')
        table.export('next.csv', f='csv', compress=False)
        # Extracting all the tables


FILENAME = "1710.05006.pdf"

# Give Password if the pdf is encrypted, Otherwise leave it.
PASSWORD = "LOKE180874"


obj = PDFEXTRACTION(FILENAME)

obj.decrypt(PASSWORD)
obj.extracting_images(2)
obj.check_directory()
obj.extracting_text(2)

obj.table_pdf()
os.remove("output.pdf")
