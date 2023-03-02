import tkinter as tk
import tkinter.filedialog as filedialog
import pandas as pd


class Ventana:
    
    def __init__(self, master):
        self.master = master
        master.title("Búsqueda de Datos")

        self.etiqueta_marca = tk.Label(master, text="Marca de vehículo:")
        self.etiqueta_marca.grid(row=0, column=0)

        self.marca_var = tk.StringVar(value='')
        self.entry_marca = tk.Entry(master, textvariable=self.marca_var)
        self.entry_marca.grid(row=0, column=1)

        self.etiqueta_resultados = tk.Label(master, text="")
        self.etiqueta_resultados.grid(row=1, column=0, columnspan=2)

        self.boton_buscar = tk.Button(master, text="Buscar", command=self.buscar_datos)
        self.boton_buscar.grid(row=0, column=2)

        self.boton_exportar = tk.Button(master, text="Exportar a Excel", state='disabled', command=self.exportar_a_excel)
        self.boton_exportar.grid(row=1, column=2)

    def buscar_datos(self):
        self.etiqueta_resultados.config(text="")
        self.boton_exportar.config(state='disabled')

        archivo = filedialog.askopenfilename(title="Seleccione el archivo",
                                             filetypes=(("Archivos de texto", "*.txt"), ("Todos los archivos", "*.*")))

        if archivo:
            try:
                marcas = self.marca_var.get().split(',')
                df = pd.read_csv(archivo, sep='|', encoding='iso-8859-1')

                # Filtrar por marcas seleccionadas
                df_filtrado = df[df['MARCA_VEHICULO'].isin(marcas)]

                # Agrupar y sumar
                resultados = df_filtrado.groupby('MARCA_VEHICULO')['CANTIDAD'].sum()

                # Mostrar resultados
                if len(resultados) == 0:
                    self.etiqueta_resultados.config(text="No se encontraron resultados.")
                else:
                    texto_resultados = ""
                    for marca, cantidad in resultados.iteritems():
                        texto_resultados += f"{marca}: {cantidad}\n"
                    self.etiqueta_resultados.config(text=texto_resultados)

                    # Activar botón de exportar
                    self.boton_exportar.config(state='normal', command=lambda: self.exportar_a_excel(resultados, df_filtrado))

            except (pd.errors.EmptyDataError, pd.errors.ParserError, KeyError):
                self.etiqueta_resultados.config(text="Error al leer el archivo.")

    def exportar_a_excel(self, resultados, df_filtrado):
        archivo_salida = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if archivo_salida:
            df_resultados = df_filtrado.groupby(['ANIO_ALZA', 'MES', 'NOMBRE_DEPARTAMENTO', 'NOMBRE_MUNICIPIO', 'MODELO_VEHICULO', 'LINEA_VEHICULO', 'TIPO_VEHICULO', 'USO_VEHICULO', 'MARCA_VEHICULO'])['CANTIDAD'].sum().reset_index()
            df_resultados.to_excel(archivo_salida, index=False)
            

root = tk.Tk()
ventana = Ventana(root)
root.mainloop()
