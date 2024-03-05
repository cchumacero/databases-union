from tkinter import *
from tkinter import filedialog
from tkinter import Menu
from tkinter import messagebox
import pandas as pd

# Lista con el path de los archivos
archivos = []

# Nombre de la columna en común
columna_comun = ''

def noHayColumna(columna_comun):
    for file in archivos:
        fileTemp = pd.read_excel(file)
        dataTemp = fileTemp.to_dict("list")
        if columna_comun not in dataTemp:
            return True

    return False
def unificar(columna_comun):
    dataframes = []
    keys = set()
    #Borrar duplicados de un archivo y remplazar los vacios por '-'
    for file in archivos:
        dfTemp = pd.read_excel(file)
        dfTemp = dfTemp.drop_duplicates(subset=[columna_comun])
        dfTemp = dfTemp.fillna('-')
        keys = keys | set(list(dfTemp.columns.values))
        dataframes.append(dfTemp)
    df = dataframes[0].set_index(columna_comun)
    d = df.to_dict("index")
    
    for i in range(1, len(dataframes), 1):
        dfTemp = dataframes[i].set_index(columna_comun)
        dTemp = dfTemp.to_dict("index")
        
        for key, item in dTemp.items():
            #si no existe el valor del campo_comun en el diccionario principal, agrego todo.
            if key not in d:
                d.update({key: item})
            #si existe, agrego los valores restantes al campo(key) ya existente
            else:
                for key2, item2 in dTemp[key].items():
                    if key2 not in d[key] or item2 != '-':
                        d[key].update({key2: item2})
    
    dFinal = {}
    keys.remove(columna_comun)
    dFinal[columna_comun] = []
    for s in keys:
        dFinal[s] = []

    for k, i in d.items():
        (dFinal[columna_comun]).append(k)
        for s in keys:
            try:
                value = d[k][s]
            except:
                value = '-'
            dFinal[s].append(value)
    
            

    print(d)
    print(dFinal)
    final = (pd.DataFrame(dFinal))
    final.to_excel('final.xlsx')






        
# Funcion que se ejecuta cuando se hace click en el botón Unificar
def click_unificar():
    error_msg = ''
    columna_comun = columna_entry.get()
    #Chequeo de si hay una columna en común
    hayError = noHayColumna(columna_comun)
    if hayError:
        error_msg = 'La columna no es válida.'
    #Chequear luego, cuando seleccionas una db vacia

    # Deberán programar ustedes algo parecido a:
    # dataframe, error_msg = funcion_de_unificar(archivos, columna_comun)
    if len(error_msg) > 0:
        messagebox.showerror('Error!', error_msg)
    else:
        unificar(columna_comun)
        messagebox.showinfo('Operación exitosa!', "Listas correctamente unificadas")
        # Guardar el archivo (deberan programar)

# Funcion que se ejecuta cuando se hace click en el botón Abrir
def openfile():
    # Cuadro de dialogo para abrir el archivo
    nuevo_archivo = filedialog.askopenfilename(initialdir = "/",title = "Seleccionar archivo",filetypes = (("Formato excel","*.xlsx"),("all files","*.*")))

    # Si se selecciono correctamente un archivo
    if len(nuevo_archivo) > 0:
        viejo_texto = archivos_label.cget("text")

        # Si es el primer archivo:
        if len(archivos) == 0:
            viejo_texto = " Archivos seleccionados:"
            columna_entry.configure(state="normal")    # Habilito el Entry de texto
            unificar_button.configure(state="normal")  # Habilito el botón unificar

        nuevo_texto =  viejo_texto + "\n - " + nuevo_archivo   # Agrego el nuevo archivo a la lista
        archivos_label.configure(text=nuevo_texto)             # Actualizo el texto del Label

        archivos.append(nuevo_archivo)  # Lo agrego a la lista

window = Tk()
window.title("Unificador de bases de datos")
window.geometry('500x200')

# Menu 'Archivo' - 'Abrir'
menu = Menu(window)
new_item = Menu(menu, tearoff=0)
new_item.add_command(label='Abrir', command=openfile)
menu.add_cascade(label='Archivo', menu=new_item)
window.config(menu=menu)

# Label con la lista de archivos
archivos_label = Label(window, text=" No se seleccionaron archivos", justify=LEFT)
archivos_label.grid(column=0, row=0)

# Entrada de texto del nombre de la columna en comun
columna_label = Label(window, text="\n Columna en común:")
columna_label.grid(column=0, row=1)
columna_entry = Entry(window,width=20, state='disabled')  # Empieza desabilitada
columna_entry.grid(column=0, row=2)

# Boton de unificar
unificar_button = Button(window, text="Unificar listas", command=click_unificar, state = 'disabled')  # Empieza desabilitado
unificar_button.grid(column=0, row=3, pady=10)

window.mainloop()