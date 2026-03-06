package lab6file;

import javax.swing.*;
import javax.swing.text.*;
import java.awt.*;
import java.awt.event.*;
import java.util.*;
import org.apache.poi.xwpf.usermodel.*;

public class GestorTablas {

    private JTextPane editorTexto;
    private List<JTable> tablasInsertadas = new ArrayList<>();

    public GestorTablas(JTextPane editorTexto) {
        this.editorTexto = editorTexto;
    }

    public void insertarTabla() {
        DialogoTabla dialogo = new DialogoTabla((JFrame) SwingUtilities.getWindowAncestor(editorTexto));
        dialogo.setVisible(true);

        if (dialogo.isAceptado()) {
            int filas    = dialogo.getFilas();
            int columnas = dialogo.getColumnas();
            crearTablaVisual(filas, columnas, null);
        }
    }

    private void crearTablaVisual(int filas, int columnas, String[][] datos) {
        DefaultTableModel modelo = new DefaultTableModel(filas, columnas);

        if (datos != null) {
            for (int f = 0; f < filas; f++) {
                for (int c = 0; c < columnas; c++) {
                    if (f < datos.length && c < datos[f].length) {
                        modelo.setValueAt(datos[f][c], f, c);
                    }
                }
            }
        }

        JTable tabla = new JTable(modelo);
        tabla.setGridColor(Color.DARK_GRAY);
        tabla.setRowHeight(28);
        tabla.setFont(new Font("Arial", Font.PLAIN, 13));
        tabla.getTableHeader().setBackground(new Color(220, 220, 220));
        tabla.setFillsViewportHeight(true);

        int anchoTotal = columnas * 100;
        int altoTotal  = (filas + 1) * 30;
        tabla.setPreferredScrollableViewportSize(new Dimension(anchoTotal, altoTotal));

        JScrollPane scroll = new JScrollPane(tabla);
        scroll.setPreferredSize(new Dimension(anchoTotal + 20, altoTotal + 20));

        editorTexto.insertComponent(scroll);
        tablasInsertadas.add(tabla);

        try {
            StyledDocument doc = editorTexto.getStyledDocument();
            doc.insertString(doc.getLength(), "\n", null);
        } catch (BadLocationException e) {
            e.printStackTrace();
        }
    }

    public void guardarTablas(XWPFDocument documento) {
        for (JTable tabla : tablasInsertadas) {
            int filas    = tabla.getRowCount();
            int columnas = tabla.getColumnCount();

            XWPFTable tablaWord = documento.createTable(filas, columnas);

            for (int f = 0; f < filas; f++) {
                XWPFTableRow fila = tablaWord.getRow(f);
                for (int c = 0; c < columnas; c++) {
                    Object valor = tabla.getValueAt(f, c);
                    String texto = (valor != null) ? valor.toString() : "";

                    XWPFTableCell celda;
                    if (c < fila.getTableCells().size()) {
                        celda = fila.getCell(c);
                    } else {
                        celda = fila.addNewTableCell();
                    }

                    celda.setText(texto);
                }
            }
        }
    }

    public void cargarTablas(XWPFDocument documento) {
        List<XWPFTable> tablasWord = documento.getTables();

        for (XWPFTable tablaWord : tablasWord) {
            int filas    = tablaWord.getRows().size();
            int columnas = (filas > 0) ? tablaWord.getRow(0).getTableCells().size() : 0;

            if (filas == 0 || columnas == 0) continue;

            String[][] datos = new String[filas][columnas];
            for (int f = 0; f < filas; f++) {
                XWPFTableRow fila = tablaWord.getRow(f);
                for (int c = 0; c < columnas; c++) {
                    if (c < fila.getTableCells().size()) {
                        datos[f][c] = fila.getCell(c).getText();
                    } else {
                        datos[f][c] = "";
                    }
                }
            }

            crearTablaVisual(filas, columnas, datos);
        }
    }

    public void limpiar() {
        tablasInsertadas.clear();
    }

    static class DialogoTabla extends JDialog {

        private JSpinner spinnerFilas;
        private JSpinner spinnerColumnas;
        private boolean aceptado = false;

        public DialogoTabla(JFrame padre) {
            super(padre, "Insertar Tabla", true);
            construirVentana();
        }

        private void construirVentana() {
            setSize(280, 180);
            setLocationRelativeTo(getParent());
            setResizable(false);

            JPanel panel = new JPanel(new GridBagLayout());
            panel.setBorder(BorderFactory.createEmptyBorder(15, 20, 10, 20));
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(6, 6, 6, 6);
            gbc.fill   = GridBagConstraints.HORIZONTAL;

            gbc.gridx = 0; gbc.gridy = 0;
            panel.add(new JLabel("Filas:"), gbc);
            gbc.gridx = 1;
            spinnerFilas = new JSpinner(new SpinnerNumberModel(3, 1, 20, 1));
            panel.add(spinnerFilas, gbc);

            gbc.gridx = 0; gbc.gridy = 1;
            panel.add(new JLabel("Columnas:"), gbc);
            gbc.gridx = 1;
            spinnerColumnas = new JSpinner(new SpinnerNumberModel(3, 1, 10, 1));
            panel.add(spinnerColumnas, gbc);

            JPanel panelBotones = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 0));
            JButton btnAceptar  = new JButton("Aceptar");
            JButton btnCancelar = new JButton("Cancelar");

            btnAceptar.setPreferredSize(new Dimension(90, 28));
            btnCancelar.setPreferredSize(new Dimension(90, 28));

            btnAceptar.addActionListener(e -> {
                aceptado = true;
                dispose();
            });

            btnCancelar.addActionListener(e -> {
                aceptado = false;
                dispose();
            });

            panelBotones.add(btnAceptar);
            panelBotones.add(btnCancelar);

            gbc.gridx = 0; gbc.gridy = 2;
            gbc.gridwidth = 2;
            panel.add(panelBotones, gbc);

            add(panel);
        }

        public boolean isAceptado()  { return aceptado; }
        public int getFilas()        { return (int) spinnerFilas.getValue(); }
        public int getColumnas()     { return (int) spinnerColumnas.getValue(); }
    }
}