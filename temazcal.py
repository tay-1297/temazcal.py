import tkinter as tk
from tkinter import ttk
import mysql.connector
import pandas as pd
import os
from reportlab.pdfgen import canvas
from tkinter import messagebox

def crear_boton(marco, texto, funcion):
    return tk.Button(marco, text=texto, command=funcion, width=20, height=2, bg="Royalblue2", fg="white", font=("Arial", 11, "bold"), activebackground="blue", activeforeground="white", relief="flat", bd=0)

def crear_titulo(marco, texto):
    lbl_titulo = tk.Label(marco, text=texto, font=("Arial", 22, "bold"), bg="white", fg="RoyalBlue4", padx=20, pady=15)
    lbl_titulo.pack(fill="x", pady=(10, 20))
    return lbl_titulo

def conectar():
    try:
        conexion1 = mysql.connector.connect(host="localhost", user="root", passwd="", database="Negocio")
        print("Conexin correcta")
    except (mysql.connector.Error) as e:
        print("ocurrio un error al conctar: ", e)
    return conexion1

def mostrar_inicio():
    limpiar_contenido()
    crear_titulo(marco_contenido, "CONTROL DE INVENTARIO")
    tk.Label(marco_contenido, text="🏬 Materias Primas", fg="RoyalBlue4", bg="white", font=("Arial", 22, "bold")).pack(pady=10)
    tk.Label(marco_contenido, text="El Temazcal", fg="RoyalBlue4", bg="white", font=("Arial", 22, "bold")).pack(pady=10)

def mostrar_productos():
    limpiar_contenido()
    def mostrarPro():
        con = conectar(); c1 = con.cursor(); c1.execute("select * from productos")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
    def insertarPro():
        con = conectar(); c1 = con.cursor()
        co = cajaC.get(); de = cajaD.get(); pr = cajaP.get()
        c1.execute("insert into productos values(%s,%s,%s)", (co, de, pr)); con.commit()
    def eliminarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get()
        c1.execute(f"DELETE FROM productos WHERE pro_codigo_k ={co}", (co)); con.commit()
        c1.execute("select * from productos")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def buscarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get()
        c1.execute(f"select * from productos WHERE pro_codigo_k ={co}")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def actualizarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get(); de = cajaD.get()
        c1.execute("UPDATE productos SET pro_descripcion=%s WHERE pro_codigo_k=%s", (de, co)); con.commit()
        c1.execute("select * from productos")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def generar_excel():
        columnas = tabla["columns"]
        datos = [tabla.item(fila)["values"] for fila in tabla.get_children()]
        if not datos:
            messagebox.showwarning("Aviso", "No hay datos para generar el reporte"); return
        df = pd.DataFrame(datos, columns=columnas)
        df.to_excel("reporte_inventario.xlsx", index=False)
        os.startfile("reporte_inventario.xlsx")
    def generar_PDF():
        pdf = canvas.Canvas("reporte_pro.pdf")
        pdf.setFont("Helvetica-Bold", 16); pdf.drawString(150, 750, "REPORTE DE PRODUCTOS")
        pdf.setFont("Helvetica-Bold", 10); pdf.drawString(50, 700, "CODIGO"); pdf.drawString(100, 700, "DESCRIPCION")
        y = 670
        for fila in tabla.get_children():
            valores = tabla.item(fila)["values"]
            pdf.setFont("Helvetica", 10); pdf.drawString(50, y, str(valores[0])); pdf.drawString(100, y, str(valores[1])); y -= 25
        pdf.save(); os.startfile("reporte_pro.pdf")
    crear_titulo(marco_contenido, "PRODUCTOS")
    tabla = ttk.Treeview(marco_contenido, columns=("codigo", "descripcion", "precio"), show="headings")
    tabla.heading("codigo", text="CODIGO"); tabla.heading("descripcion", text="DESCRIPCIÒN"); tabla.heading("precio", text="PRECIO")
    tabla.column("codigo", width=100); tabla.column("descripcion", width=100); tabla.column("precio", width=100)
    tabla.config(height=10); tabla.place(x=50, y=80)
    botonI = tk.Button(marco_contenido, text="Mostrar", width=12, command=mostrarPro); botonI.place(x=50, y=50)
    botonS = tk.Button(marco_contenido, text="Salir", width=12, command=marco_contenido.destroy); botonS.place(x=450, y=320)
    botonIN = tk.Button(marco_contenido, text="Insertar", width=12, command=insertarPro); botonIN.place(x=450, y=360)
    botonE = tk.Button(marco_contenido, text="Eliminar", width=12, command=eliminarPro); botonE.place(x=450, y=400)
    botonB = tk.Button(marco_contenido, text="Buscar", width=12, command=buscarPro); botonB.place(x=450, y=440)
    botonA = tk.Button(marco_contenido, text="Actualizar", width=12, command=actualizarPro); botonA.place(x=450, y=480)
    botonRep1=tk.Button(marco_contenido,text="Reporte Excel",width=12,command=generar_excel); botonRep1.place(x=450 , y=50)
    botonRep2=tk.Button(marco_contenido,text="Reporte PDF",width=12,command=generar_PDF); botonRep2.place(x=450 , y=100)
    lbl_Codigo = tk.Label(marco_contenido, text="Codigo").place(x=20, y=360)
    cajaC = tk.Entry(marco_contenido, width=25); cajaC.place(x=70, y=360)
    lbl_Descripcion = tk.Label(marco_contenido, text="Descripcion").place(x=20, y=390)
    cajaD = tk.Entry(marco_contenido, width=25); cajaD.place(x=90, y=390)
    lbl_precio = tk.Label(marco_contenido, text="PRECIO").place(x=20, y=420)
    cajaP = tk.Entry(marco_contenido, width=25); cajaP.place(x=90, y=420)

def mostrar_exixtencias():
    limpiar_contenido()
    def mostrarPro():
        con = conectar(); c1 = con.cursor(); c1.execute("select * from existencias")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
    def insertarPro():
        con = conectar(); c1 = con.cursor()
        co = cajaC.get(); lo = cajaD.get(); can = cajaCan.get()
        c1.execute("insert into existencias values(%s,%s,%s)", (co, lo, can)); con.commit()
    def eliminarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get()
        c1.execute(f"DELETE FROM existencias WHERE pro_codigo_k ={co}", (co)); con.commit()
        c1.execute("select * from existencias")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def buscarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get()
        c1.execute(f"select * from existencias WHERE pro_codigo_k ={co}")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def actualizarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get(); de = cajaD.get()
        c1.execute("UPDATE existencias SET pro_descripcion=%s WHERE pro_codigo_k=%s", (de, co)); con.commit()
        c1.execute("select * from productos")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def generar_excel():
        columnas = tabla["columns"]
        datos = [tabla.item(fila)["values"] for fila in tabla.get_children()]
        if not datos:
            messagebox.showwarning("Aviso", "No hay datos para generar el reporte"); return
        df = pd.DataFrame(datos, columns=columnas)
        df.to_excel("reporte_inventario.xlsx", index=False)
        os.startfile("reporte_inventario.xlsx")
    crear_titulo(marco_contenido, "EXISTENCIA")
    tabla = ttk.Treeview(marco_contenido, columns=("codigo", "lote", "cantidad"), show="headings")
    tabla.heading("codigo", text="CODIGO"); tabla.heading("lote", text="LOTE"); tabla.heading("cantidad", text="CANTIDAD")
    tabla.column("codigo", width=100); tabla.column("lote", width=100); tabla.column("cantidad", width=100)
    tabla.config(height=10); tabla.place(x=50, y=90)
    botonI = tk.Button(marco_contenido, text="Mostrar", width=12, command=mostrarPro); botonI.place(x=450, y=280)
    botonS = tk.Button(marco_contenido, text="Salir", width=12, command=marco_contenido.destroy); botonS.place(x=450, y=320)
    botonIN = tk.Button(marco_contenido, text="Insertar", width=12, command=insertarPro); botonIN.place(x=450, y=360)
    botonE = tk.Button(marco_contenido, text="Eliminar", width=12, command=eliminarPro); botonE.place(x=450, y=400)
    botonB = tk.Button(marco_contenido, text="Buscar", width=12, command=buscarPro); botonB.place(x=450, y=440)
    botonA = tk.Button(marco_contenido, text="Actualizar", width=12, command=actualizarPro); botonA.place(x=450, y=480)
    botonRep1=tk.Button(marco_contenido,text="Reporte Excel",width=12,command=generar_excel); botonRep1.place(x=450 , y=500)
    lbl_Codigo = tk.Label(marco_contenido, text="Codigo").place(x=20, y=360)
    cajaC = tk.Entry(marco_contenido, width=25); cajaC.place(x=70, y=360)
    lbl_lote = tk.Label(marco_contenido, text="LOTE").place(x=20, y=390)
    cajaD = tk.Entry(marco_contenido, width=25); cajaD.place(x=70, y=390)
    lbl_cantidad = tk.Label(marco_contenido, text="CANTIDAD").place(x=20, y=420)
    cajaCan = tk.Entry(marco_contenido, width=25); cajaCan.place(x=90, y=420)

def mostrar_movimientos():
    limpiar_contenido()
    def mostrarPro():
        con = conectar(); c1 = con.cursor(); c1.execute("select * from movimientos")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
    def insertarPro():
        con = conectar(); c1 = con.cursor()
        co = cajaC.get(); mo = cajaT.get(); fe = cajaD.get(); cod = cajaCo.get(); lo = cajaL.get(); cos = cajaCos.get(); pre = cajaPre.get(); can = cajaCa.get(); to = cajaTo.get()
        if to == "" or to is None:
            try: to = float(pre) * float(can)
            except: to = 0
        c1.execute("insert into movimientos values(%s,%s,%s,%s,%s,%s,%s,%s,%s)", (co, mo, fe, cod, lo, cos, pre, can, to))
        con.commit()
        messagebox.showinfo("Listo", "Movimiento guardado")
    def eliminarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get()
        c1.execute(f"DELETE FROM movimientos WHERE mov_codigo_k ={co}", (co)); con.commit()
        c1.execute("select * from movimientos")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def buscarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get()
        c1.execute(f"select * from movimientos WHERE mov_codigo_k ={co}")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def actualizarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get(); de = cajaD.get()
        c1.execute("UPDATE movimientos SET pro_descripcion=%s WHERE mov_codigo_k=%s", (de, co)); con.commit()
        c1.execute("select * from movimientos")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    # --- LOS 3 QUE PEDISTE: EXCEL + PDF + TICKET ---
    def generar_excel():
        try:
            if not tabla.get_children():
                messagebox.showwarning("Aviso", "Primero dale a Mostrar"); return
            import datetime
            nombre = f"reporte_movimientos_{datetime.datetime.now().strftime('%H%M%S')}.xlsx"
            columnas = ["CODIGO","MOVIMIENTO","FECHA","CODIGO PROD","LOTE","COSTO","PRECIO","CANTIDAD","TOTAL"]
            datos = []
            for fila in tabla.get_children():
                valores = list(tabla.item(fila)["values"])
                valores = [0 if v is None or v == "None" else v for v in valores]
                # Si total es 0, lo calcula
                try:
                    if float(valores[8]) == 0:
                        valores[8] = float(valores[6]) * float(valores[7])
                except: pass
                datos.append(valores)
            df = pd.DataFrame(datos, columns=columnas)
            df.to_excel(nombre, index=False)
            messagebox.showinfo("Listo", f"Ya se creó {nombre}")
            os.startfile(nombre)
        except Exception as e:
            messagebox.showerror("Error", str(e))
    def generar_PDF():
        try:
            if not tabla.get_children():
                messagebox.showwarning("Aviso", "Primero dale a Mostrar"); return
            import datetime
            nombre = f"reporte_movimientos_{datetime.datetime.now().strftime('%H%M%S')}.pdf"
            pdf = canvas.Canvas(nombre)
            pdf.setFont("Helvetica-Bold", 16); pdf.drawString(150, 750, "REPORTE DE MOVIMIENTOS")
            pdf.setFont("Helvetica-Bold", 8)
            pdf.drawString(20, 700, "CODIGO"); pdf.drawString(70, 700, "MOV"); pdf.drawString(110, 700, "FECHA"); pdf.drawString(180, 700, "COD PROD"); pdf.drawString(250, 700, "LOTE"); pdf.drawString(300, 700, "COSTO"); pdf.drawString(350, 700, "PRECIO"); pdf.drawString(400, 700, "CANT"); pdf.drawString(450, 700, "TOTAL")
            y = 670
            for fila in tabla.get_children():
                valores = tabla.item(fila)["values"]
                valores = [0 if v is None or v == "None" else v for v in valores]
                try:
                    if float(valores[8]) == 0:
                        valores[8] = float(valores[6]) * float(valores[7])
                except: pass
                pdf.setFont("Helvetica", 8)
                pdf.drawString(20, y, str(valores[0])); pdf.drawString(70, y, str(valores[1])); pdf.drawString(110, y, str(valores[2])); pdf.drawString(180, y, str(valores[3])); pdf.drawString(250, y, str(valores[4])); pdf.drawString(300, y, str(valores[5])); pdf.drawString(350, y, str(valores[6])); pdf.drawString(400, y, str(valores[7])); pdf.drawString(450, y, str(valores[8]))
                y -= 25
                if y < 50:
                    pdf.showPage(); y = 750
            pdf.save()
            messagebox.showinfo("Listo", f"Ya se creó {nombre}")
            os.startfile(nombre)
        except Exception as e:
            messagebox.showerror("Error", str(e))
    def generar_ticket():
        try:
            seleccion = tabla.selection()
            if not seleccion:
                messagebox.showwarning("Aviso", "Primero selecciona una fila azul y luego Ticket"); return
            valores = tabla.item(seleccion[0])["values"]
            valores = [0 if v is None or v == "None" else v for v in valores]
            codigo = valores[0]; mov = valores[1]; fecha = valores[2]; pro_codigo = valores[3]; lote = valores[4]; costo = valores[5]; precio = valores[6]; cantidad = valores[7]; total = valores[8]
            if total == 0:
                try: total = float(precio) * float(cantidad)
                except: total = 0
            import datetime
            nombre = f"ticket_venta_{datetime.datetime.now().strftime('%H%M%S')}.pdf"
            pdf = canvas.Canvas(nombre, pagesize=(220, 400))
            pdf.setFont("Helvetica-Bold", 12); pdf.drawString(50, 370, "EL TEMAZCAL")
            pdf.setFont("Helvetica", 9)
            pdf.drawString(20, 340, "Materias Primas")
            pdf.drawString(20, 320, f"Ticket No: {codigo}")
            pdf.drawString(20, 305, f"Fecha: {fecha}")
            pdf.drawString(20, 290, f"Movimiento: {mov}")
            pdf.drawString(20, 275, f"Codigo Prod: {pro_codigo}")
            pdf.drawString(20, 260, f"Lote: {lote}")
            pdf.drawString(20, 245, f"Costo: ${costo}")
            pdf.drawString(20, 230, f"Precio: ${precio}")
            pdf.drawString(20, 215, f"Cantidad: {cantidad}")
            pdf.drawString(20, 195, "----------------------")
            pdf.setFont("Helvetica-Bold", 11); pdf.drawString(20, 180, f"TOTAL: ${total}")
            pdf.setFont("Helvetica", 9); pdf.drawString(20, 150, "Gracias por su compra")
            pdf.save()
            messagebox.showinfo("Listo", f"Ya se creó {nombre}")
            os.startfile(nombre)
        except Exception as e:
            messagebox.showerror("Error", str(e))
    crear_titulo(marco_contenido, "MOVIMIENTOS")
    tabla = ttk.Treeview(marco_contenido, columns=("codigo", "movimiento", "fecha", "pro_Codigo", "lote", "costo","precio", "cantidad", "total"), show="headings")
    tabla.heading("codigo", text="CODIGO"); tabla.heading("movimiento", text="MOVIMIENTO"); tabla.heading("fecha", text="FECHA"); tabla.heading("pro_Codigo", text="Codigo"); tabla.heading("lote", text="LOTE"); tabla.heading("costo", text="COSTO"); tabla.heading("precio", text="PRECIO"); tabla.heading("cantidad", text="CANTIDAD"); tabla.heading("total", text="TOTAL")
    tabla.column("codigo", width=80); tabla.column("movimiento", width=80); tabla.column("fecha", width=90); tabla.column("pro_Codigo", width=80); tabla.column("lote", width=80); tabla.column("costo", width=80); tabla.column("precio", width=80); tabla.column("cantidad", width=80); tabla.column("total", width=80)
    tabla.config(height=10); tabla.place(x=40, y=250)
    botonI = tk.Button(marco_contenido, text="Mostrar", width=12, command=mostrarPro); botonI.place(x=100, y=100)
    botonS = tk.Button(marco_contenido, text="🏠 Salir", width=12, command=mostrar_inicio); botonS.place(x=200, y=100)
    botonIN = tk.Button(marco_contenido, text="Insertar", width=12, command=insertarPro); botonIN.place(x=300, y=100)
    botonE = tk.Button(marco_contenido, text="Eliminar", width=12, command=eliminarPro); botonE.place(x=400, y=100)
    botonB = tk.Button(marco_contenido, text="Buscar", width=12, command=buscarPro); botonB.place(x=500, y=100)
    botonA = tk.Button(marco_contenido, text="Actualizar", width=12, command=actualizarPro); botonA.place(x=600, y=100)
    # LOS 3 BOTONES JUNTOS
    botonRep1=tk.Button(marco_contenido,text="Reporte Excel",width=12,command=generar_excel); botonRep1.place(x=700 , y=100)
    botonRep2=tk.Button(marco_contenido,text="Reporte PDF",width=12,command=generar_PDF); botonRep2.place(x=800 , y=100)
    botonTicket=tk.Button(marco_contenido,text="Ticket",width=12,command=generar_ticket, bg="#E91E63", fg="white", font=("Arial", 10, "bold")); botonTicket.place(x=900 , y=100)
    lbl_Codigo = tk.Label(marco_contenido, text="Codigo 1").place(x=870, y=360)
    cajaC = tk.Entry(marco_contenido, width=25); cajaC.place(x=950, y=360)
    lbl_tipo = tk.Label(marco_contenido, text="Movimiento").place(x=870, y=390)
    cajaT = ttk.Combobox(marco_contenido, values=["ENT", "SAL"], state="readonly", width=17); cajaT.place(x=950, y=390)
    lbl_Descripcion = tk.Label(marco_contenido, text="Fecha").place(x=870, y=420)
    cajaD = tk.Entry(marco_contenido, width=25); cajaD.place(x=950, y=420)
    lbl_Codigo = tk.Label(marco_contenido, text="Codigo 2").place(x=870, y=450)
    cajaCo = tk.Entry(marco_contenido, width=25); cajaCo.place(x=950, y=450)
    lbl_Codigo = tk.Label(marco_contenido, text="Lote").place(x=870, y=480)
    cajaL = tk.Entry(marco_contenido, width=25); cajaL.place(x=950, y=480)
    lbl_Codigo = tk.Label(marco_contenido, text="Costo").place(x=870, y=510)
    cajaCos = tk.Entry(marco_contenido, width=25); cajaCos.place(x=950, y=510)
    lbl_Codigo = tk.Label(marco_contenido, text="Precio").place(x=870, y=540)
    cajaPre = tk.Entry(marco_contenido, width=25); cajaPre.place(x=950, y=540)
    lbl_Codigo = tk.Label(marco_contenido, text="Cantidad").place(x=870, y=570)
    cajaCa = tk.Entry(marco_contenido, width=25); cajaCa.place(x=950, y=570)
    lbl_Codigo = tk.Label(marco_contenido, text="Total").place(x=870, y=600)
    cajaTo = tk.Entry(marco_contenido, width=25); cajaTo.place(x=950, y=600)

def mostrar_provedores():
    limpiar_contenido()
    def mostrarPro():
        con = conectar(); c1 = con.cursor(); c1.execute("select * from provedores")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
    def insertarPro():
        con = conectar(); c1 = con.cursor()
        co = cajaC.get(); de = cajaD.get(); te = cajaP.get()
        c1.execute("insert into provedores values(%s,%s,%s)", (co, de, te)); con.commit()
    def eliminarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get()
        c1.execute(f"DELETE FROM provedores WHERE prov_id ={co}", (co)); con.commit()
        c1.execute("select * from provedores")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def buscarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get()
        c1.execute(f"select * from provedores WHERE prov_id ={co}")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    def actualizarPro():
        con = conectar(); c1 = con.cursor(); co = cajaC.get(); de = cajaD.get()
        c1.execute("UPDATE provedores SET prov_nombre=%s WHERE prov_id=%s", (de, co)); con.commit()
        c1.execute("select * from provedores")
        tabla.delete(*tabla.get_children())
        for fila in c1:
            tabla.insert(parent="", index=tk.END, values=fila)
        con.close()
    crear_titulo(marco_contenido, "PROVEDORES")
    botonI = tk.Button(marco_contenido, text="Mostrar", width=12, command=mostrarPro); botonI.place(x=50, y=50)
    botonS = tk.Button(marco_contenido, text="Salir", width=12, command=mostrar_inicio); botonS.place(x=450, y=320)
    botonIN = tk.Button(marco_contenido, text="Insertar", width=12, command=insertarPro); botonIN.place(x=450, y=360)
    botonE = tk.Button(marco_contenido, text="Eliminar", width=12, command=eliminarPro); botonE.place(x=450, y=400)
    botonB = tk.Button(marco_contenido, text="Buscar", width=12, command=buscarPro); botonB.place(x=450, y=440)
    botonA = tk.Button(marco_contenido, text="Actualizar", width=12, command=actualizarPro); botonA.place(x=450, y=480)
    lbl_Codigo = tk.Label(marco_contenido, text="ID Provedor").place(x=20, y=360)
    cajaC = tk.Entry(marco_contenido, width=25); cajaC.place(x=90, y=360)
    lbl_Descripcion = tk.Label(marco_contenido, text="Provedor").place(x=20, y=390)
    cajaD = tk.Entry(marco_contenido, width=25); cajaD.place(x=90, y=390)
    lbl_precio = tk.Label(marco_contenido, text="Telefono").place(x=20, y=420)
    cajaP = tk.Entry(marco_contenido, width=25); cajaP.place(x=90, y=420)
    tabla = ttk.Treeview(marco_contenido, columns=("id provedor", "provedor", "telefono"), show="headings")
    tabla.heading("id provedor", text="ID PROVEDOR"); tabla.heading("provedor", text="PROVEDOR"); tabla.heading("telefono", text="TELEFONO")
    tabla.column("id provedor", width=100); tabla.column("provedor", width=100); tabla.column("telefono", width=100)
    tabla.config(height=10); tabla.place(x=50, y=80)

def limpiar_contenido():
    for widget in marco_contenido.winfo_children():
        widget.destroy()

def mostrar_principal():
    global marco_contenido, marco_menu
    ventana = tk.Tk(); ventana.title("Sistema de Control de Inventario"); ventana.state("zoomed")
    marco_menu = tk.Frame(ventana, width=200, bg="Royalblue3"); marco_menu.pack(side=tk.LEFT, fill=tk.Y)
    marco_contenido = tk.Frame(ventana, bg="alice blue"); marco_contenido.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH)
    marco_encabezado = tk.Frame(marco_contenido, bg="white", height=80); marco_encabezado.pack(fill=tk.X)
    crear_boton(marco_menu, "🏠 INICIO", mostrar_inicio).pack(pady=10)
    crear_boton(marco_menu, "🛒 PRODUCTOS", mostrar_productos).pack(pady=10)
    crear_boton(marco_menu, "🍞 EXISTENCIAS", mostrar_exixtencias).pack(pady=10)
    crear_boton(marco_menu, "💵 MOVIMIENTOS", mostrar_movimientos).pack(pady=10)
    crear_boton(marco_menu, "🚛 PROVEDORES", mostrar_provedores).pack(pady=10)
    crear_boton(marco_menu, "🚶 SALIR", ventana.destroy).pack(pady=10)
    mostrar_inicio(); ventana.mainloop()

mostrar_principal()