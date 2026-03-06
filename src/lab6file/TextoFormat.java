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

public void tachado() {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
if (inicio == fin) return;

boolean tieneTachado = verificarAtributo(inicio, fin, StyleConstants.StrikeThrough);
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setStrikeThrough(atributo, !tieneTachado);
documento.setCharacterAttributes(inicio, fin - inicio, atributo, false);
areaTexto.requestFocus();
}

public void superindice() {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
if (inicio == fin) return;

boolean tieneSuperindice = verificarAtributo(inicio, fin, StyleConstants.Superscript);
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setSuperscript(atributo, !tieneSuperindice);
if (!tieneSuperindice) StyleConstants.setSubscript(atributo, false);
documento.setCharacterAttributes(inicio, fin - inicio, atributo, false);
areaTexto.requestFocus();
}

public void subindice() {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
if (inicio == fin) return;

boolean tieneSubindice = verificarAtributo(inicio, fin, StyleConstants.Subscript);
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setSubscript(atributo, !tieneSubindice);
if (!tieneSubindice) StyleConstants.setSuperscript(atributo, false);
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

public void colorFondo(Component ventana) {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();

if (inicio == fin) {
JOptionPane.showMessageDialog(ventana, "Selecciona texto primero");
return;
}

Color colorElegido = JColorChooser.showDialog(ventana, "Elige color de resaltado", Color.YELLOW);

if (colorElegido != null) {
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setBackground(atributo, colorElegido);
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

public void alinearIzquierda() {
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setAlignment(atributo, StyleConstants.ALIGN_LEFT);
aplicarParrafo(atributo);
}

public void alinearCentro() {
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setAlignment(atributo, StyleConstants.ALIGN_CENTER);
aplicarParrafo(atributo);
}

public void alinearDerecha() {
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setAlignment(atributo, StyleConstants.ALIGN_RIGHT);
aplicarParrafo(atributo);
}

public void justificado() {
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setAlignment(atributo, StyleConstants.ALIGN_JUSTIFIED);
aplicarParrafo(atributo);
}

public void aumentarSangria() {
SimpleAttributeSet atributo = new SimpleAttributeSet();
int inicio = areaTexto.getSelectionStart();
Element parrafo = documento.getParagraphElement(inicio);
float sangriaActual = StyleConstants.getLeftIndent(parrafo.getAttributes());
StyleConstants.setLeftIndent(atributo, sangriaActual + 20);
aplicarParrafo(atributo);
}

public void reducirSangria() {
SimpleAttributeSet atributo = new SimpleAttributeSet();
int inicio = areaTexto.getSelectionStart();
Element parrafo = documento.getParagraphElement(inicio);
float sangriaActual = StyleConstants.getLeftIndent(parrafo.getAttributes());
float nuevaSangria = Math.max(0, sangriaActual - 20);
StyleConstants.setLeftIndent(atributo, nuevaSangria);
aplicarParrafo(atributo);
}

public void interlineado(float espacio) {
SimpleAttributeSet atributo = new SimpleAttributeSet();
StyleConstants.setLineSpacing(atributo, espacio);
aplicarParrafo(atributo);
}

private void aplicarParrafo(SimpleAttributeSet atributo) {
int inicio = areaTexto.getSelectionStart();
int fin = areaTexto.getSelectionEnd();
documento.setParagraphAttributes(inicio, fin - inicio, atributo, false);
areaTexto.requestFocus();
}

private boolean verificarAtributo(int inicio, int fin, Object key) {
for (int pos = inicio; pos < fin; pos++) {
Element elemento = documento.getCharacterElement(pos);
AttributeSet atributos = elemento.getAttributes();

boolean valor = false;
if (key == StyleConstants.Bold) valor = StyleConstants.isBold(atributos);
else if (key == StyleConstants.Italic) valor = StyleConstants.isItalic(atributos);
else if (key == StyleConstants.Underline) valor = StyleConstants.isUnderline(atributos);
else if (key == StyleConstants.StrikeThrough) valor = StyleConstants.isStrikeThrough(atributos);
else if (key == StyleConstants.Superscript) valor = StyleConstants.isSuperscript(atributos);
else if (key == StyleConstants.Subscript) valor = StyleConstants.isSubscript(atributos);

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

public boolean isTachado() {
int posicion = areaTexto.getCaretPosition();
return StyleConstants.isStrikeThrough(documento.getCharacterElement(posicion).getAttributes());
}
}

