package lab6file;

import org.apache.poi.xwpf.usermodel.*;
import javax.swing.*;
import javax.swing.text.*;
import java.awt.Color;
import java.io.*;

public class GestorArchivos{

    public static void guardar(JTextPane textPane,String ruta){
        guardar(textPane,ruta,null);
    }

    public static void guardar(JTextPane textPane,String ruta,GestorTablas gestorTablas){
        try{
            StyledDocument doc=textPane.getStyledDocument();
            String textoCompleto=textPane.getText();
            int size=textoCompleto.length();

            XWPFDocument documento=new XWPFDocument();
            XWPFParagraph parrafo=documento.createParagraph();

            int i=0;
            while(i<size){
                AttributeSet atributos=doc.getCharacterElement(i).getAttributes();

                int j=i+1;
                while(j<size){
                    AttributeSet attrsJ=doc.getCharacterElement(j).getAttributes();
                    if(!mismoFormato(atributos,attrsJ))break;
                    j++;
                }

                String fragmento=textoCompleto.substring(i,j);
                String familia=StyleConstants.getFontFamily(atributos);
                int tamano=StyleConstants.getFontSize(atributos);
                boolean negrita=StyleConstants.isBold(atributos);
                boolean cursiva=StyleConstants.isItalic(atributos);
                boolean subray=StyleConstants.isUnderline(atributos);
                Color color=StyleConstants.getForeground(atributos);

                XWPFRun run=parrafo.createRun();
                run.setText(fragmento);
                run.setFontFamily(familia);
                run.setFontSize(tamano);
                run.setBold(negrita);
                run.setItalic(cursiva);
                run.setUnderline(subray?UnderlinePatterns.SINGLE:UnderlinePatterns.NONE);

                String hex=String.format("%02X%02X%02X",
                        color.getRed(),color.getGreen(),color.getBlue());
                run.setColor(hex);

                i=j;
            }

            if(gestorTablas!=null){
                gestorTablas.guardarTablas(documento);
            }

            ByteArrayOutputStream baos=new ByteArrayOutputStream();
            documento.write(baos);
            byte[] contenidoPoi=baos.toByteArray();

            RandomAccessFile rarch=new RandomAccessFile(ruta,"rw");
            rarch.setLength(0);
            rarch.seek(0);
            rarch.write(contenidoPoi);
            rarch.close();

            documento.close();
            JOptionPane.showMessageDialog(null,"Archivo guardado correctamente.");

        }catch(Exception e){
            JOptionPane.showMessageDialog(null,"Error al guardar: "+e.getMessage());
            e.printStackTrace();
        }
    }

    public static void abrir(JTextPane textPane,String ruta){
        abrir(textPane,ruta,null);
    }

    public static void abrir(JTextPane textPane,String ruta,GestorTablas gestorTablas){
        try{
            RandomAccessFile rarch=new RandomAccessFile(ruta,"r");

            if(rarch.length()==0){
                rarch.close();
                JOptionPane.showMessageDialog(null,"El archivo está vacío.");
                return;
            }

            rarch.seek(0);
            byte[] contenido=new byte[(int)rarch.length()];
            rarch.readFully(contenido);
            rarch.close();

            XWPFDocument documento=new XWPFDocument(new ByteArrayInputStream(contenido));

            StyledDocument doc=textPane.getStyledDocument();
            doc.remove(0,doc.getLength());

            for(XWPFParagraph parrafo:documento.getParagraphs()){
                for(XWPFRun run:parrafo.getRuns()){
                    String texto=run.getText(0);
                    if(texto==null||texto.isEmpty())continue;

                    String familia=run.getFontFamily()!=null?run.getFontFamily():"Arial";
                    int tamano=run.getFontSize()>0?run.getFontSize():12;

                    boolean negrita=run.isBold();
                    boolean cursiva=run.isItalic();
                    boolean subray=run.getUnderline()!=UnderlinePatterns.NONE;

                    Color color=Color.BLACK;
                    String colorHex=run.getColor();
                    if(colorHex!=null&&!colorHex.equalsIgnoreCase("auto")){
                        color=Color.decode("#"+colorHex);
                    }

                    SimpleAttributeSet attrs=new SimpleAttributeSet();
                    StyleConstants.setFontFamily(attrs,familia);
                    StyleConstants.setFontSize(attrs,tamano);
                    StyleConstants.setBold(attrs,negrita);
                    StyleConstants.setItalic(attrs,cursiva);
                    StyleConstants.setUnderline(attrs,subray);
                    StyleConstants.setForeground(attrs,color);

                    doc.insertString(doc.getLength(),texto,attrs);
                }
                doc.insertString(doc.getLength(),"\n",null);
            }

            if(gestorTablas!=null){
                gestorTablas.cargarTablas(documento);
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