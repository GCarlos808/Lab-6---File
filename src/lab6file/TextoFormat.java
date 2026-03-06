/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package lab6file;
import javax.swing.*;
import javax.swing.text.*;
import java.awt.*;
/**
 *
 * @author riche
 */
public class TextoFormat {

JTextPane areaTexto;
StyledDocument documento;

public TextoFormat(JTextPane areaTexto) {
this.areaTexto = areaTexto;
this.documento = areaTexto.getStyledDocument();
}

public void negrita() {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
if (inicio == fin) return;

boolean tieneNegrita = verificarAtributo(inicio, fin, StyleConstants.Bold);
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setBold(atributo, !tieneNegrita);
documento.setCharacterAttributes(inicio, fin - inicio, atributo, false);
areaTexto.requestFocus();
}

public void cursiva() {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
if (inicio == fin) return;

boolean tieneCursiva = verificarAtributo(inicio, fin, StyleConstants.Italic);
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setItalic(atributo, !tieneCursiva);
documento.setCharacterAttributes(inicio, fin - inicio, atributo, false);
areaTexto.requestFocus();
}

public void subrayado() {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
if (inicio == fin) return;

boolean tieneSubrayado = verificarAtributo(inicio, fin, StyleConstants.Underline);
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setUnderline(atributo, !tieneSubrayado);
documento.setCharacterAttributes(inicio, fin - inicio, atributo, false);
areaTexto.requestFocus();
}

public void color(Component ventana) {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();

if (inicio == fin) {
JOptionPane.showMessageDialog(ventana, "Selecciona texto primero");
return;
}

Color colorActual = getColorActual(inicio);
Color colorElegido = JColorChooser.showDialog(ventana, "Elige un color", colorActual);

if (colorElegido != null) {
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setForeground(atributo, colorElegido);
documento.setCharacterAttributes(inicio, fin - inicio, atributo, false);
}
areaTexto.requestFocus();
}

public void fuente(String nombre) {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
if (inicio == fin) return;

SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setFontFamily(atributo, nombre);
documento.setCharacterAttributes(inicio, fin - inicio, atributo, false);
areaTexto.requestFocus();
}

public void tamano(int numero) {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
if (inicio == fin) return;

SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setFontSize(atributo, numero);
documento.setCharacterAttributes(inicio, fin - inicio, atributo, false);
areaTexto.requestFocus();
}

public void tamano(String numero) {
try {
tamano(Integer.parseInt(numero.trim()));
} catch (NumberFormatException e) {
System.out.println("tamaño invalido");
}
}

private boolean verificarAtributo(int inicio, int fin, Object key) {
for (int pos = inicio; pos < fin; pos++) {
Element elemento = documento.getCharacterElement(pos);
AttributeSet atributos = elemento.getAttributes();

boolean valor = false;
if (key == StyleConstants.Bold) valor = StyleConstants.isBold(atributos);
else if (key == StyleConstants.Italic) valor = StyleConstants.isItalic(atributos);
else if (key == StyleConstants.Underline) valor = StyleConstants.isUnderline(atributos);

if (!valor) return false;
}
return true;
}

private Color getColorActual(int posicion) {
Element elemento = documento.getCharacterElement(posicion);
Color color = StyleConstants.getForeground(elemento.getAttributes());
return (color != null) ? color : Color.BLACK;
}

public String getFuente() {
int posicion = areaTexto.getCaretPosition();
return StyleConstants.getFontFamily(documento.getCharacterElement(posicion).getAttributes());
}

public int getTamano() {
int posicion = areaTexto.getCaretPosition();
return StyleConstants.getFontSize(documento.getCharacterElement(posicion).getAttributes());
}

public boolean isNegrita() {
int posicion = areaTexto.getCaretPosition();
return StyleConstants.isBold(documento.getCharacterElement(posicion).getAttributes());
}

public boolean isCursiva() {
int posicion = areaTexto.getCaretPosition();
return StyleConstants.isItalic(documento.getCharacterElement(posicion).getAttributes());
}

public boolean isSubrayado() {
int posicion = areaTexto.getCaretPosition();
return StyleConstants.isUnderline(documento.getCharacterElement(posicion).getAttributes());
}
}