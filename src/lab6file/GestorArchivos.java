package lab6file;

/**
 * Persona 3 — Gestión de archivos .docx
 * Por ahora guarda/abre con RandomAccessFile en formato binario propio.
 */

import org.apache.poi.xwpf.usermodel.*;
import javax.swing.*;
import javax.swing.text.*;
import java.awt.Color;
import java.io.*;

public class GestorArchivos {

    public static void guardar(JTextPane textPane,String ruta){
        try{
            StyledDocument doc= textPane.getStyledDocument();
            String textoCompleto= textPane.getText();
            int largo=textoCompleto.length();

            XWPFDocument documento= new XWPFDocument();
            XWPFParagraph parrafo=documento.createParagraph();

            ByteArrayOutputStream baos= new ByteArrayOutputStream();
            DataOutputStream dos= new DataOutputStream(baos);

            int i=0;
            while(i<largo){
                AttributeSet attrs=doc.getCharacterElement(i).getAttributes();

                int j=i+1;
                while(j<largo){
                    AttributeSet attrsJ=doc.getCharacterElement(j).getAttributes();
                    if(!mismoFormato(attrs,attrsJ))break;
                    j++;
                }

                String fragmento=textoCompleto.substring(i,j);

                String familia=StyleConstants.getFontFamily(attrs);
                int tamanio=StyleConstants.getFontSize(attrs);
                boolean negrita=StyleConstants.isBold(attrs);
                boolean cursiva=StyleConstants.isItalic(attrs);
                boolean subray=StyleConstants.isUnderline(attrs);
                Color color=StyleConstants.getForeground(attrs);

                byte[] bytesTexto=fragmento.getBytes("UTF-8");
                byte[] bytesFont=familia.getBytes("UTF-8");

                dos.writeInt(bytesTexto.length);
                dos.write(bytesTexto);
                dos.writeInt(bytesFont.length);
                dos.write(bytesFont);
                dos.writeInt(tamanio);
                dos.writeBoolean(negrita);
                dos.writeBoolean(cursiva);
                dos.writeBoolean(subray);
                dos.writeInt(color.getRed());
                dos.writeInt(color.getGreen());
                dos.writeInt(color.getBlue());

                XWPFRun run=parrafo.createRun();
                run.setText(fragmento);
                run.setFontFamily(familia);
                run.setFontSize(tamanio);
                run.setBold(negrita);
                run.setItalic(cursiva);
                if(subray)run.setUnderline(UnderlinePatterns.SINGLE);
                else run.setUnderline(UnderlinePatterns.NONE);

                String hex=String.format("%02X%02X%02X",
                        color.getRed(),color.getGreen(),color.getBlue());
                run.setColor(hex);

                i=j;
            }

            dos.close();
            byte[] contenido=baos.toByteArray();

            RandomAccessFile raf=new RandomAccessFile(ruta,"rw");
            raf.setLength(0);
            raf.seek(0);
            raf.write(contenido);
            raf.close();

            ByteArrayOutputStream baosPoi=new ByteArrayOutputStream();
            documento.write(baosPoi);
            byte[] contenidoPoi=baosPoi.toByteArray();

            raf=new RandomAccessFile(ruta,"rw");
            raf.setLength(0);
            raf.seek(0);
            raf.write(contenidoPoi);
            raf.close();

            documento.close();

            JOptionPane.showMessageDialog(null,"Archivo guardado correctamente.");

        }catch(Exception e){
            JOptionPane.showMessageDialog(null,"Error al guardar: "+e.getMessage());
            e.printStackTrace();
        }
    }

    public static void abrir(JTextPane textPane,String ruta){
        try{
            RandomAccessFile raf=new RandomAccessFile(ruta,"r");

            if(raf.length()==0){
                raf.close();
                JOptionPane.showMessageDialog(null,"El archivo está vacío.");
                return;
            }

            raf.seek(0);
            byte[] contenido=new byte[(int)raf.length()];
            raf.readFully(contenido);
            raf.close();

            ByteArrayInputStream bais=new ByteArrayInputStream(contenido);
            XWPFDocument documento=new XWPFDocument(bais);

            StyledDocument doc=textPane.getStyledDocument();
            doc.remove(0,doc.getLength());

            DataInputStream dis=new DataInputStream(new ByteArrayInputStream(contenido));

            while(dis.available()>0){
                int largoTexto=dis.readInt();
                byte[] bytesTexto=new byte[largoTexto];
                dis.readFully(bytesTexto);
                String texto=new String(bytesTexto,"UTF-8");

                int largoFont=dis.readInt();
                byte[] bytesFont=new byte[largoFont];
                dis.readFully(bytesFont);
                String familia=new String(bytesFont,"UTF-8");

                int tamanio=dis.readInt();
                boolean negrita=dis.readBoolean();
                boolean cursiva=dis.readBoolean();
                boolean subray=dis.readBoolean();

                int r=dis.readInt();
                int g=dis.readInt();
                int b=dis.readInt();
                Color color=new Color(r,g,b);

                SimpleAttributeSet attrs=new SimpleAttributeSet();
                StyleConstants.setFontFamily(attrs,familia);
                StyleConstants.setFontSize(attrs,tamanio);
                StyleConstants.setBold(attrs,negrita);
                StyleConstants.setItalic(attrs,cursiva);
                StyleConstants.setUnderline(attrs,subray);
                StyleConstants.setForeground(attrs,color);

                doc.insertString(doc.getLength(),texto,attrs);
            }

            dis.close();

            for(XWPFParagraph parrafo:documento.getParagraphs()){
                for(XWPFRun run:parrafo.getRuns()){
                    String texto=run.getText(0);
                    if(texto==null||texto.isEmpty())continue;

                    String familia=run.getFontFamily()!=null?run.getFontFamily():"Arial";
                    int tamanio=run.getFontSize()>0?run.getFontSize():12;

                    boolean negrita=run.isBold();
                    boolean cursiva=run.isItalic();
                    boolean subray=run.getUnderline()!=UnderlinePatterns.NONE;

                    Color color=Color.BLACK;
                    String colorHex=run.getColor();
                    if(colorHex!=null&&!colorHex.equalsIgnoreCase("auto"))
                        color=Color.decode("#"+colorHex);

                    SimpleAttributeSet attrs=new SimpleAttributeSet();
                    StyleConstants.setFontFamily(attrs,familia);
                    StyleConstants.setFontSize(attrs,tamanio);
                    StyleConstants.setBold(attrs,negrita);
                    StyleConstants.setItalic(attrs,cursiva);
                    StyleConstants.setUnderline(attrs,subray);
                    StyleConstants.setForeground(attrs,color);

                    doc.insertString(doc.getLength(),texto,attrs);
                }
                doc.insertString(doc.getLength(),"\n",null);
            }

            documento.close();

            JOptionPane.showMessageDialog(null,"Archivo abierto correctamente.");

        }catch(Exception e){
            JOptionPane.showMessageDialog(null,"Error al abrir: "+e.getMessage());
            e.printStackTrace();
        }
    }

    private static boolean mismoFormato(AttributeSet a,AttributeSet b){
        return StyleConstants.getFontFamily(a).equals(StyleConstants.getFontFamily(b))
                && StyleConstants.getFontSize(a)==StyleConstants.getFontSize(b)
                && StyleConstants.isBold(a)==StyleConstants.isBold(b)
                && StyleConstants.isItalic(a)==StyleConstants.isItalic(b)
                && StyleConstants.isUnderline(a)==StyleConstants.isUnderline(b)
                && StyleConstants.getForeground(a).equals(StyleConstants.getForeground(b));
    }
}