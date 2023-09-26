import openpyxl
import os

# Lista para almacenar los datos a copiar
datas = []

class CopyFile():
    def __init__(self, **kwargs):
        # Crear un nuevo libro de trabajo
        self.wb = openpyxl.Workbook()

        # Crear el nombre del archivo de salida
        filename = kwargs["filename"] + ".xlsx"

        # Abrir el archivo Excel existente para copiar datos
        file_to_copy = openpyxl.load_workbook("Lista-de-clientes-con-nombre-y-direccion.xlsx")
        dataframe = file_to_copy.active


        # Crear una nueva hoja en el libro de trabajo de salida
        new_file = self.wb.active
        new_file.title = kwargs["title"]

        # Copiar los datos de la hoja de cálculo original a la nueva hoja
        for row in range(1, dataframe.max_row):
            data_for_process = [row, ]
            for data in dataframe.iter_cols(1, dataframe.max_column):
                data_for_process.append(data[row].value)
            datas.append(data_for_process)



        filas_names = []
        rows = [int(x) for x in kwargs["row"].split(",")]
        for i in range(len(rows)):
            nombre_fila = input(f"Escribe el nombre de la fila #{i + 1}: ")
            filas_names.append(nombre_fila)
        print("Creando archivo 75%..................")


        for index, value in enumerate(rows):
            z = 2
            print(index)
            if value == 1:
                for x in range(len(datas)):
                    # Copiar los datos de la fila especificada
                    copy = datas[x][value]
                    # Asignar los datos copiados a la nueva hoja de cálculo
                    if index == 0:
                        new_file[f"A{z}"] = copy
                    elif index == 1:
                        new_file[f"B{z}"] = copy
                    elif index == 2:
                        new_file[f"C{z}"] = copy
                    elif index == 3:
                        new_file[f"D{z}"] = copy
                    elif index == 4:
                        new_file[f"E{z}"] = copy         
                    z += 1
            if value == 2:
                for x in range(len(datas)):
                    # Copiar los datos de la fila especificada
                    copy = datas[x][value]
                    # Asignar los datos copiados a la nueva hoja de cálculo
                    if index == 0:
                        new_file[f"A{z}"] = copy
                    elif index == 1:
                        new_file[f"B{z}"] = copy
                    elif index == 2:
                        new_file[f"C{z}"] = copy
                    elif index == 3:
                        new_file[f"D{z}"] = copy
                    elif index == 4:
                        new_file[f"E{z}"] = copy
                    z += 1
            if value == 3:
                for x in range(len(datas)):
                    # Copiar los datos de la fila especificada
                    copy = datas[x][3]
                    # Asignar los datos copiados a la nueva hoja de cálculo
                    if index == 0:
                        new_file[f"A{z}"] = copy
                    elif index == 1:
                        new_file[f"B{z}"] = copy
                    elif index == 2:
                        new_file[f"C{z}"] = copy
                    elif index == 3:
                        new_file[f"D{z}"] = copy
                    elif index == 4:
                        new_file[f"E{z}"] = copy
                    z += 1
            if value == 4:
                for x in range(len(datas)):
                    # Copiar los datos de la fila especificada
                    copy = datas[x][value]
                    # Asignar los datos copiados a la nueva hoja de cálculo
                    if index == 0:
                        new_file[f"A{z}"] = copy
                    elif index == 1:
                        new_file[f"B{z}"] = copy
                    elif index == 2:
                        new_file[f"C{z}"] = copy
                    elif index == 3:
                        new_file[f"D{z}"] = copy
                    elif index == 4:
                        new_file[f"E{z}"] = copy
                    z += 1
            if value == 5:
                for x in range(len(datas)):
                    # Copiar los datos de la fila especificada
                    copy = datas[x][value]
                    # Asignar los datos copiados a la nueva hoja de cálculo
                    if index == 0:
                        new_file[f"A{z}"] = copy
                    elif index == 1:
                        new_file[f"B{z}"] = copy
                    elif index == 2:
                        new_file[f"C{z}"] = copy
                    elif index == 3:
                        new_file[f"D{z}"] = copy
                    elif index == 4:
                        new_file[f"E{z}"] = copy
                    z += 1
                

        for i in range(len(filas_names)):
            if i == 0:
                new_file[f"A1"] = filas_names[i]
            if i == 1:
                new_file[f"B1"] = filas_names[i]
                i -= 1
            if i == 2:
                new_file[f"C1"] = filas_names[i]
                i -= 1
            if i == 3:
                new_file[f"D1"] = filas_names[i]
                i -= 1
            if i == 4:
                new_file[f"E1"] = filas_names[i]
                i -= 1

        # Guardar el archivo Excel de salida
        self.wb.save(filename)

        # Cerrar el archivo Excel de salida
        self.wb.close()

        print("Archivo creado con éxito")

if __name__ == "__main__":
    train = False
    while train == False:
        filename = input("Nombre del archivo (solo el nombre, sin extensiones): ")
        title = input("Ponle un título a la hoja: ")
        row = (input("Escribe el número de fila a copiar (si son múltiples, sepáralos con comas sin espacios): "))
        archivo = os.path.exists(filename + ".xlsx")
        if archivo == True:
            print("El archivo ya existe, por favor ingresa otro nombre")
        else:
            train = True
    else:
        print("Creando archivo...")
        copy_obj = CopyFile(row=row, title=title, filename=filename)
