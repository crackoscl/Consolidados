import xlwings as xw
import datetime
import locale
import win32api
import win32con
import os.path as ruta

class Consolidado:

    def __init__(self):
        self.wb = xw.Book("consolidado.xlsm")
        self.datos = self.wb.sheets("datos")
        self.temp = self.wb.sheets("temp")
        self.folios = self.wb.sheets("folios")
        self.cantidad_pallets = int()
        self.contenedor = ""
        self.path = ""
        self.especie = ""

    def copy_cells_visble_filter(self):
        self.temp.api.Select()
        self.temp.api.Cells.Clear()
        self.datos.api.Select()
        self.datos.api.UsedRange.SpecialCells(12).Copy()
        self.temp.select()
        self.temp.api.Range("A1").Select()
        self.temp.api.Paste()

    def set_especies(self)->str:
        self.especie = " / ".join(set([cell.value for cell in self.temp.range("R2").expand("down")]))
 
    def set_quantity_pallets(self)->int:
        self.cantidad_pallets = len(set([cell.value for cell in self.temp.range("N2").expand("down")]))

    def set_path(self)->str:
        self.contenedor = self.temp["C2"].value
        nombre_mes = self.temp["A2"].value.strftime("%B")
        self.path = r"\\backu\Respaldo\Respaldo Despacho\Informes Consolidado\Fotografias Consolidado 23-24\{} 23\{}\{}".format(
        nombre_mes,self.temp["A2"].value.strftime("%d-%m-%Y"),self.contenedor
        )

    def temp_bcm(self,especie)->str:
        temps = {
            "CEREZAS": ["5% 15 CBM", "-1.0"],
            "CIRUELAS": ["5% 15 CBM", "-0.5"],
            "NECTARINES": ["5% 15 CBM", "-0.5"],
            "PERAS": ["10% 35 CBM", "-1.0"],
            "UVAS": ["0% 0 CBM", "-0.5"],
            "KIWIS": ["5% 15 CBM", "-0.5"],
        }

        return temps[especie][0], temps[especie][1]
    

    def delete_pictures_range(self,range_address, plano):
        picture_range = plano.range(range_address)
        for picture in plano.pictures:
            if (
                picture.left >= picture_range.left
                and picture.top >= picture_range.top
                and picture.left + picture.width <= picture_range.left + picture_range.width
                and picture.top + picture.height <= picture_range.top + picture_range.height
            ):
                picture.delete()

    def rgbToInt(self,rgb)->int:
        colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
        return colorInt

    def add_border_range(self,range_address,plano):
        shapes_range = plano.range(range_address)
        for shape in plano.shapes:
            if (
                shape.left >= shapes_range.left
                and shape.top >= shapes_range.top
                and shape.left + shape.width <= shapes_range.left + shapes_range.width
                and shape.top + shape.height <= shapes_range.top + shapes_range.height
            ):
                shape.api.Line.ForeColor.RGB = self.rgbToInt((0, 255, 0))
                shape.api.Line.Weight = 1.2

    def copy_folios(self,planob):
        planob["B41"].value = self.folios["B2"].value
        planob["C41"].value = self.folios["B14"].value
        planob["B43"].value = self.folios["B3"].value
        planob["C43"].value = self.folios["B15"].value
        planob["B45"].value = self.folios["B4"].value
        planob["C45"].value = self.folios["B16"].value
        planob["B45"].value = self.folios["B4"].value
        planob["C45"].value = self.folios["B16"].value
        planob["B47"].value = self.folios["B5"].value
        planob["C47"].value = self.folios["B17"].value
        planob["B49"].value = self.folios["B6"].value
        planob["C49"].value = self.folios["B18"].value
        planob["B51"].value = self.folios["B7"].value
        planob["C51"].value = self.folios["B19"].value
        planob["B53"].value = self.folios["B8"].value
        planob["C53"].value = self.folios["B20"].value
        planob["B55"].value = self.folios["B9"].value
        planob["C55"].value = self.folios["B21"].value
        planob["B57"].value = self.folios["B10"].value
        planob["C57"].value = self.folios["B22"].value
        planob["B59"].value = self.folios["B11"].value
        planob["C59"].value = self.folios["B23"].value


    def add_pictures_plano(self,fotos,plano):
        for key,value in fotos.items():
            if ruta.exists(self.path + "\\"+ f"{key}"):
                plano.pictures.add(
                    self.path
                    + "\\"
                        + f"{key}",
                        left=plano.range(f"{value[0]}").left,
                        top=plano.range(f"{value[1]}").top,
                        width=380,
                        height=272,
                    )
            else:
                win32api.MessageBox(
                        self.wb.app.hwnd,
                        f" no se puede encontrar el archivo {key}",
                        "Error",
                        win32con.MB_ICONERROR,
                    )
                 
    def generate_sheet_a(self):
        if self.cantidad_pallets ==23:
            planoa = self.wb.sheets("23Pallets(A)")
            planoa.api.Visible = True
            self.wb.sheets("20Pallets(A)").api.Visible = False
            self.wb.sheets("21Pallets(A)").api.Visible = False
        elif self.cantidad_pallets== 21:
            planoa = self.wb.sheets("21Pallets(A)")
            planoa.api.Visible = True
            self.wb.sheets("20Pallets(A)").api.Visible = False
            self.wb.sheets("23Pallets(A)").api.Visible = False
        elif 0 < self.cantidad_pallets <= 20:
            planoa = self.wb.sheets("20Pallets(A)")
            planoa.api.Visible = True
            self.wb.sheets("23Pallets(A)").api.Visible = False
            self.wb.sheets("21Pallets(A)").api.Visible = False

        planoa["C7"].value = self.temp["O2"].value
        planoa["C8"].value = self.temp["A2"].value
        planoa["C9"].value = self.especie
        planoa["C10"].value = self.temp["B2"].value
        planoa["C11"].value = self.contenedor
        planoa["C13"].value = self.temp["G2"].value.strip()
        planoa["C14"].value = self.temp["I2"].value.strip()
        planoa["C15"].value = self.temp["D2"].value.strip()
        planoa["H8"].value = self.temp["Q2"].value
        planoa["H9"].value = self.temp["M2"].value
        planoa["H10"].value = self.temp["L2"].value
        planoa["H11"].value = self.cantidad_pallets
        planoa["H13"].value = self.temp["H2"].value
        planoa["H14"].value = self.temp["J2"].value
        planoa.select()
        self.delete_pictures_range("A29:I42", planoa)
        fotos = {
            "tem1.JPG":["B30","C30"],
            "tem2.JPG":["G30","H30"],
        }
        self.add_pictures_plano(fotos,planoa)

        self.add_border_range("A29:I42", planoa)

    def generate_sheet_b(self):
        if  self.cantidad_pallets == 23:
            planob = self.wb.sheets("23Pallets(B)")
            planob.api.Visible = True
            self.wb.sheets("21Pallets(B)").api.Visible = False
            self.wb.sheets("20Pallets(B)").api.Visible = False
            self.planob["B61"].value = self.folios["B12"].value
            planob["C61"].value = self.folios["B24"].value
            planob["B63"].value = self.folios["B13"].value

        elif self.cantidad_pallets == 21:
            planob = self.wb.sheets("21Pallets(B)")
            planob.api.Visible = True
            self.wb.sheets("20Pallets(B)").api.Visible = False
            self.wb.sheets("23Pallets(B)").api.Visible = False
            planob["B61"].value = self.folios["B12"].value

        elif 0 < self.cantidad_pallets <= 20:
            print("aqui")
            planob = self.wb.sheets("20Pallets(B)")
            planob.api.Visible = True
            self.wb.sheets("23Pallets(B)").api.Visible = False
            self.wb.sheets("21Pallets(B)").api.Visible = False

        planob["C11"].value = "LINDEROS"
        planob["C12"].value = self.temp["A2"].value
        planob["C13"].value = self.especie
        planob["C14"].value = self.temp["B2"].value
        planob["C15"].value = self.contenedor
        planob["C17"].value = self.temp["Q2"].value
        planob["C18"].value = self.temp["O2"].value
        planob["C20"].value = self.temp["G2"].value.strip()
        planob["C21"].value = self.temp["I2"].value.strip()
        planob["C22"].value = self.temp["D2"].value.strip()
        planob["G11"].value = self.temp["M2"].value
        planob["G12"].value = self.temp["L2"].value
        planob["G13"].value = self.temp["T2"].value
        planob["G14"].value = self.temp["E2"].value + " / " + self.temp["F2"].value
        planob["G15"].value = self.temp["K2"].value
        hora_despacho_TimeSerial = int(self.temp["V2"].value)
        hora_despacho = datetime.datetime.combine(
        datetime.date.today(),
        datetime.time(
            hora_despacho_TimeSerial // 10000,
            (hora_despacho_TimeSerial // 100) % 100,
            hora_despacho_TimeSerial % 100,
            ),
        )
        hora_inicio_carga = hora_despacho - datetime.timedelta(minutes=30)
        planob["G16"].value = f"{hora_inicio_carga.hour}:{hora_inicio_carga.minute}"
        planob["G17"].value = f"{hora_despacho.hour}:{hora_despacho.minute}"
        planob["G18"].value = self.cantidad_pallets
        planob["G20"].value = self.temp["H2"].value
        planob["G21"].value = self.temp["J2"].value
        bcm, temp = self.temp_bcm(self.especie)
        planob["G27"].value = bcm
        planob["C29"].value = temp
        self.copy_folios(planob)
        planob.select()

        self.delete_pictures_range("A68:G103", planob)

        planob.range("G54:G61").value = chr(0x2713)
  
        fotos = {
            "sigla.JPG":["B69","C69"],
            "seteo.JPG":["F69","G69"],
            "lampa.JPG":["B81","C81"],
            "piso.JPG":["F81","G81"],
            "cierre.JPG":["B92","C92"],
        }
        self.add_pictures_plano(fotos,planob)
        
        self.add_border_range("A67:H103", planob)
      

def generar_consolidado():
    locale.setlocale(locale.LC_ALL, "")
    consolidado = Consolidado()
    consolidado.copy_cells_visble_filter()
    consolidado.set_path()
    consolidado.set_especies()
    consolidado.set_quantity_pallets()
    if ruta.exists(consolidado.path):
        consolidado.generate_sheet_a()
        consolidado.generate_sheet_b()
        win32api.MessageBox(
                    consolidado.wb.app.hwnd,
                    "LLenado de Datos Finalazado !!!",
                    "Info",
                    win32con.MB_ICONINFORMATION,
                )
    else:
        win32api.MessageBox(
                consolidado.wb.app.hwnd,
                f"carpeta nombre {consolidado.contenedor} no existe o esta mal escrita",
                "Error",
                win32con.MB_ICONERROR,
            )