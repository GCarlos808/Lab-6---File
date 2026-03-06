package lab6file;

import javax.swing.*;
import javax.swing.table.*;
import javax.swing.text.*;
import java.awt.*;
import java.awt.event.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;

public class GestorTablas {

    private JTextPane editorTexto;
    private List<JTable> tablasInsertadas = new ArrayList<>();
    private List<Map<String, Color>> coloresPorTabla = new ArrayList<>();

    public GestorTablas(JTextPane editorTexto) {
        this.editorTexto = editorTexto;
    }

    public void insertarTabla() {
        DialogoTabla dialogo = new DialogoTabla((JFrame) SwingUtilities.getWindowAncestor(editorTexto));
        dialogo.setVisible(true);

        if (dialogo.isAceptado()) {
            int filas    = dialogo.getFilas();
            int columnas = dialogo.getColumnas();
            crearTablaVisual(filas, columnas, null, null, null);
        }
    }

    private void crearTablaVisual(int filas, int columnas, String[] encabezados, String[][] datos, Map<String, Color> coloresGuardados) {
        if (encabezados == null) {
            encabezados = new String[columnas];
            for (int c = 0; c < columnas; c++) {
                encabezados[c] = "Columna " + (c + 1);
            }
        }

        final String[] encabezadosFinal = encabezados;

        DefaultTableModel modelo = new DefaultTableModel(encabezadosFinal, filas) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return true;
            }
        };

        if (datos != null) {
            for (int f = 0; f < filas; f++) {
                for (int c = 0; c < columnas; c++) {
                    if (f < datos.length && c < datos[f].length) {
                        modelo.setValueAt(datos[f][c], f, c);
                    }
                }
            }
        }

        Map<String, Color> coloresCelda = (coloresGuardados != null) ? coloresGuardados : new HashMap<>();
        coloresPorTabla.add(coloresCelda);

        JTable tabla = new JTable(modelo);
        tabla.setGridColor(Color.DARK_GRAY);
        tabla.setRowHeight(28);
        tabla.setFont(new Font("Arial", Font.PLAIN, 13));
        tabla.getTableHeader().setBackground(new Color(220, 220, 220));
        tabla.setFillsViewportHeight(true);
        tabla.getTableHeader().setReorderingAllowed(false);

        tabla.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable t, Object value,
                    boolean isSelected, boolean hasFocus, int row, int col) {
                Component c = super.getTableCellRendererComponent(t, value, isSelected, hasFocus, row, col);
                String key = row + "," + col;
                if (!isSelected) {
                    c.setBackground(coloresCelda.getOrDefault(key, Color.WHITE));
                }
                return c;
            }
        });

        tabla.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                if (SwingUtilities.isRightMouseButton(e)) {
                    int fila = tabla.rowAtPoint(e.getPoint());
                    int col  = tabla.columnAtPoint(e.getPoint());
                    if (fila == -1 || col == -1) return;

                    String key = fila + "," + col;
                    Color actual = coloresCelda.getOrDefault(key, Color.WHITE);
                    Color elegido = JColorChooser.showDialog(editorTexto, "Color de celda", actual);
                    if (elegido != null) {
                        coloresCelda.put(key, elegido);
                        tabla.repaint();
                    }
                }
            }
        });

        tabla.getTableHeader().addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) {
                    int col = tabla.columnAtPoint(e.getPoint());
                    String nuevoNombre = JOptionPane.showInputDialog("Nombre de columna:",
                            tabla.getColumnName(col));
                    if (nuevoNombre != null && !nuevoNombre.trim().isEmpty()) {
                        tabla.getColumnModel().getColumn(col).setHeaderValue(nuevoNombre);
                        tabla.getTableHeader().repaint();
                    }
                }
            }
        });

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
        for (int t = 0; t < tablasInsertadas.size(); t++) {
            JTable tabla = tablasInsertadas.get(t);
            Map<String, Color> coloresCelda = coloresPorTabla.get(t);

            int filas    = tabla.getRowCount();
            int columnas = tabla.getColumnCount();

            XWPFTable tablaWord = documento.createTable(filas + 1, columnas);

            XWPFTableRow filaEncabezado = tablaWord.getRow(0);
            for (int c = 0; c < columnas; c++) {
                String encabezado = tabla.getColumnName(c);
                XWPFTableCell celda;
                if (c < filaEncabezado.getTableCells().size()) {
                    celda = filaEncabezado.getCell(c);
                } else {
                    celda = filaEncabezado.addNewTableCell();
                }
                celda.setText(encabezado);
            }

            for (int f = 0; f < filas; f++) {
                XWPFTableRow fila = tablaWord.getRow(f + 1);
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

                    String key = f + "," + c;
                    if (coloresCelda.containsKey(key)) {
                        Color color = coloresCelda.get(key);
                        String hex = String.format("%02X%02X%02X", color.getRed(), color.getGreen(), color.getBlue());
                        CTShd shd = celda.getCTTc().addNewTcPr().addNewShd();
                        shd.setVal(STShd.CLEAR);
                        shd.setFill(hex);
                    }
                }
            }
        }
    }

    public void cargarTablas(XWPFDocument documento) {
        List<XWPFTable> tablasWord = documento.getTables();

        for (XWPFTable tablaWord : tablasWord) {
            int totalFilas = tablaWord.getRows().size();
            int columnas   = (totalFilas > 0) ? tablaWord.getRow(0).getTableCells().size() : 0;

            if (totalFilas == 0 || columnas == 0) continue;

            String[] encabezados = new String[columnas];
            XWPFTableRow filaEncabezado = tablaWord.getRow(0);
            for (int c = 0; c < columnas; c++) {
                encabezados[c] = filaEncabezado.getCell(c).getText();
            }

            int filasDatos = totalFilas - 1;
            String[][] datos = new String[filasDatos][columnas];
            Map<String, Color> coloresCelda = new HashMap<>();

            for (int f = 0; f < filasDatos; f++) {
                XWPFTableRow fila = tablaWord.getRow(f + 1);
                for (int c = 0; c < columnas; c++) {
                    if (c < fila.getTableCells().size()) {
                        XWPFTableCell celda = fila.getCell(c);
                        datos[f][c] = celda.getText();

                        if (celda.getCTTc().getTcPr() != null && celda.getCTTc().getTcPr().getShd() != null) {
                            String fill = celda.getCTTc().getTcPr().getShd().xgetFill().getStringValue();
                            if (fill != null && !fill.equalsIgnoreCase("auto") && fill.length() == 6) {
                                try {
                                    Color color = Color.decode("#" + fill);
                                    coloresCelda.put(f + "," + c, color);
                                } catch (Exception ignored) {}
                            }
                        }
                    } else {
                        datos[f][c] = "";
                    }
                }
            }

            crearTablaVisual(filasDatos, columnas, encabezados, datos, coloresCelda);
        }
    }

    public void limpiar() {
        tablasInsertadas.clear();
        coloresPorTabla.clear();
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

        public boolean isAceptado() { return aceptado; }
        public int getFilas()       { return (int) spinnerFilas.getValue(); }
        public int getColumnas()    { return (int) spinnerColumnas.getValue(); }
    }
}