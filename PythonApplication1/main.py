# -*- coding: utf-8 -*-
import os
import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QMainWindow, QVBoxLayout, QMessageBox
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle
from ast import Lambda
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QGraphicsDropShadowEffect
from PyQt5.QtCore import QPropertyAnimation,QEasingCurve
from PyQt5.QtGui import QColor
from PyQt5.uic import loadUi


class rooti(QMainWindow):
    def __init__(self):
        super(rooti,self).__init__()
        loadUi('D:/ESPE/Practicas INEN/interfaz2/PythonApplication1/interfaz2.ui',self)
        self.bt_menu.clicked.connect(self.mover_menu)
        self.bt_restart.hide()
        #botones de control
        self.bt_equi1.clicked.connect(self.equipo1)
        #self.bt_equi2.clicked.connect(self.equipo2)
        self.bt_load1.clicked.connect(self.load1)
        #self.bt_load2.clicked.connect(self.load2)
        self.bt_filter1.clicked.connect(self.filter1)
        #self.bt_filter2.clicked.connect(self.filter2)
        self.bt_save1.clicked.connect(self.save1)
        #self.bt_save2.clicked.connect(self.save2)
        #botones
        self.bt_maxi.clicked.connect(self.maximizar)
        self.bt_restart.clicked.connect(self.normal)
        self.bt_mini.clicked.connect(self.minimizar)
        self.bt_close.clicked.connect(lambda: self.close())
        
        #eliminar la ventana del main
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setWindowOpacity(1)
        #SizeGrip
        self.gripSize=10
        self.grip=QtWidgets.QSizeGrip(self)
        self.grip.resize(self.gripSize,self.gripSize)
        #mover ventana
        self.frame_top.mouseMoveEvent=self.mover_ventana
        #conexion botones
        self.bt_equi1.clicked.connect(lambda:self.stackedWidget.setCurrentWidget(self.page_equipo1))
        self.bt_equi2.clicked.connect(lambda:self.stackedWidget.setCurrentWidget(self.page_equipo2))
        
        

    

    def minimizar(self):
        self.showMinimized()
        
    def normal(self):
        self.showNormal()
        self.bt_restart.hide()
        self.bt_maxi.show()
    
    def maximizar(self): 
        self.showMaximized()
        self.bt_restart.show()
        self.bt_maxi.hide()
        
    def mousePressEvent(self,event):
        self.click_position=event.globalPos()
    
    def mover_ventana(self,event):
        if self.isMaximized()==False:
            if event.buttons()==QtCore.Qt.LeftButton:
                self.move(self.pos()+event.globalPos()-self.click_position)
                self.click_position=event.globalPos()
                event.accept()
        if event.globalPos().y()<=10:
            self.showMaximized()
            self.bt_maxi.hide()
            self.bt_restart.show()
        else:
            self.showNormal()
            self.bt_maxi.show()
            self.bt_restart.hide()
                           

 
    def mover_menu(self):
        if True:
            width=self.frame_control.width()
            normal=0
            if width==0:
                extender=200
            else:
                extender=normal
            self.animacion=QPropertyAnimation(self.frame_control,b'minimumWidth')
            self.animacion.setDuration(300)
            self.animacion.setStartValue(width)
            self.animacion.setEndValue(extender)
            self.animacion.setEasingCurve(QtCore.QEasingCurve.InOutQuart)
            self.animacion.start()
    
    def equipo1(self):
        self.bt_load1.setEnabled(True)
        
    def load1(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        archivo, _ = QFileDialog.getOpenFileName(self, 'Buscador de csv', '', 'Archivos CSV (*.csv);;Todos los archivos (*)', options=options)
        if archivo:
            self.label_4.setText("Archivo cargado")
            self.archivo_seleccionado = archivo
            self.bt_filter1.setEnabled(True)
        else:
            self.label_4.setText('No se seleccion\xF3 ning\xFAn archivo.')
    
    def filter1(self):
        if self.archivo_seleccionado:
            try:
                df = pd.read_csv(self.archivo_seleccionado)
                columnas = df.columns
                tabla = df[['Abs (Corr)1', 'Abs (Corr)2', 'Abs (Corr)3']]
                self.tabla = tabla
                self.label_4.setText("Archivo Filtrado")
                self.bt_save1.setEnabled(True)
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Error al filtrar la tabla: {str(e)}')
        else:
            QMessageBox.warning(self, 'Advertencia', 'No se seleccion\u00F3 ning\xFAn archivo.')
    
    
    def save1(self):
        if hasattr(self, 'tabla'):
            try:
                nombre_archivo = os.path.splitext(os.path.basename(self.archivo_seleccionado))[0]
                nuevo_nombre_excel = f"Filtrado_{nombre_archivo}.xlsx"
                options = QFileDialog.Options()
                options |= QFileDialog.ReadOnly
                ruta_guardado, _ = QFileDialog.getSaveFileName(self, 'Guardar en Excel', nuevo_nombre_excel, 'Archivos Excel (*.xlsx);;Todos los archivos (*)', options=options)
                
                if ruta_guardado:
                    self.tabla.to_excel(ruta_guardado, index=False)
                    QMessageBox.information(self, 'Informaci\u00F3n', f'Archivo Excel guardado en: {ruta_guardado}')
                else:
                    QMessageBox.warning(self, 'Advertencia', 'No se seleccion\u00F3 una ruta de guardado.')
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Error al guardar en Excel: {str(e)}')
        else:
            QMessageBox.warning(self, 'Advertencia', 'No se seleccion\u00F3 ninguna tabla.')





    


    
    
            
            
          

        
if __name__=="__main__":
    app=QApplication(sys.argv)
    mi_app=rooti()
    mi_app.show()
    sys.exit(app.exec_())