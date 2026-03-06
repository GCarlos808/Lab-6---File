package lab6file;

import javax.swing.*;
import javax.swing.text.*;
import javax.swing.text.rtf.RTFEditorKit;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.*;
import java.util.List;

public class GUI extends JFrame {
    
    private final JTextPane textPane;
    private final StyledDocument doc;
    private final JLabel statusBar;
    private File currentFile;
    
    private JComboBox<String> fontNameBox;
    private JComboBox<Integer> fontSizeBox;
    private JButton btnBold, btnItalic, btnUnderline;
    private JButton btnAlignLeft, btnAlignCenter, btnAlignRight;
    private JButton btnColor;
    private JButton btnBgColor;
    
    private Color currentFontColor = Color.BLACK;
    private Color currentBgColor   = null;
    
    private static final String[] FONTS = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
    private static final Integer[] SIZES = {8,9,10,11,12,14,16,18,20,22,24,26,28,32,36,40,48,56,64,72,96,144,190,240,300};
    
    public GUI() {
        super("Editor de Texto DOCX");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(1100, 750);
        setLocationRelativeTo(null);
        
        textPane = new JTextPane();
        textPane.setFont(new Font("Arial", Font.PLAIN, 12));
        doc = textPane.getStyledDocument();
        
        statusBar = new JLabel("  Listo");
        statusBar.setBorder(BorderFactory.createEtchedBorder());
        
        setJMenuBar(buildMenuBar());
        add(buildToolBar(), BorderLayout.NORTH);
        add(new JScrollPane(textPane), BorderLayout.CENTER);
        add(statusBar, BorderLayout.SOUTH);
        
        textPane.addCaretListener(e -> syncToolbarToSelection());
        
        setVisible(true);
    }
    
    private JMenuBar buildMenuBar() {
        JMenuBar mb = new JMenuBar();
        
        JMenu mFile = new JMenu("Archivo");
        addItem(mFile, "Nuevo",           KeyEvent.VK_N, e -> actionNew());
        addItem(mFile, "Abrir DOCX...",   KeyEvent.VK_O, e -> actionOpen());
        mFile.addSeparator();
        addItem(mFile, "Guardar",         KeyEvent.VK_S, e -> actionSave());
        addItem(mFile, "Guardar como...", 0,             e -> actionSaveAs());
        mFile.addSeparator();
        addItem(mFile, "Salir",           KeyEvent.VK_Q, e -> System.exit(0));
        mb.add(mFile);
        
        JMenu mFormat = new JMenu("Formato");
        addItem(mFormat, "Fuente...",  0, e -> showFontDialog());
        addItem(mFormat, "Color de texto...", 0, e -> pickFontColor());
        addItem(mFormat, "Color de fondo...", 0, e -> pickBgColor());
        mb.add(mFormat);
        
        JMenu mTable = new JMenu("Tabla");
        addItem(mTable, "Insertar tabla...", 0, e -> showTableDialog());
        mb.add(mTable);
        
        JMenu mHelp = new JMenu("Ayuda");
        addItem(mHelp, "Acerca de...", 0, e ->
            JOptionPane.showMessageDialog(this,
                "Editor de Texto DOCX\nSoporta: fuentes, colores, tamaños, tablas\nFormato .docx con OpenXML",
                "Acerca de", JOptionPane.INFORMATION_MESSAGE));
        mb.add(mHelp);
        
        return mb;
    }
    
    private void addItem(JMenu menu, String title, int key, ActionListener al) {
        JMenuItem item = new JMenuItem(title);
        if (key != 0) item.setAccelerator(KeyStroke.getKeyStroke(key, InputEvent.CTRL_DOWN_MASK));
        item.addActionListener(al);
        menu.add(item);
    }
    
    private JToolBar buildToolBar() {
        JToolBar tb = new JToolBar();
        tb.setFloatable(false);
        tb.setBackground(new Color(245, 245, 250));
        
        fontNameBox = new JComboBox<>(FONTS);
        fontNameBox.setSelectedItem("Arial");
        fontNameBox.setMaximumSize(new Dimension(200, 30));
        fontNameBox.setRenderer(new FontComboRenderer());
        fontNameBox.addActionListener(e -> applyFontName());
        tb.add(new JLabel(" Fuente: "));
        tb.add(fontNameBox);
        
        tb.addSeparator();
        
        fontSizeBox = new JComboBox<>(SIZES);
        fontSizeBox.setSelectedItem(12);
        fontSizeBox.setMaximumSize(new Dimension(70, 30));
        fontSizeBox.setEditable(true);
        fontSizeBox.addActionListener(e -> applyFontSize());
        tb.add(new JLabel(" Tamaño: "));
        tb.add(fontSizeBox);
        
        tb.addSeparator();
        
        btnBold      = makeStyleButton("B",  "Negrita (Ctrl+B)",  Font.BOLD,   e -> toggleBold());
        btnItalic    = makeStyleButton("I",  "Cursiva (Ctrl+I)",  Font.ITALIC, e -> toggleItalic());
        btnUnderline = makeStyleButton("U",  "Subrayado (Ctrl+U)",0,           e -> toggleUnderline());
        stylizeButton(btnBold,      new Font("Serif", Font.BOLD, 14));
        stylizeButton(btnItalic,    new Font("Serif", Font.ITALIC, 14));
        stylizeButton(btnUnderline, new Font("Serif", Font.PLAIN, 14));
        tb.add(btnBold);
        tb.add(btnItalic);
        tb.add(btnUnderline);
        
        textPane.getInputMap().put(KeyStroke.getKeyStroke("control B"), "bold");
        textPane.getActionMap().put("bold", new AbstractAction() 
                
        {@Override public void actionPerformed(ActionEvent e) { toggleBold(); } });
        textPane.getInputMap().put(KeyStroke.getKeyStroke("control I"), "italic");
        textPane.getActionMap().put("italic", new AbstractAction()
                
        {@Override public void actionPerformed(ActionEvent e) { toggleItalic(); } });
        textPane.getInputMap().put(KeyStroke.getKeyStroke("control U"), "underline");
        textPane.getActionMap().put("underline", new AbstractAction()
                
        {@Override public void actionPerformed(ActionEvent e) { toggleUnderline(); } });
        
        tb.addSeparator();
        
        btnAlignLeft   = makeIconButton(" ", "Alinear izquierda", e -> applyAlignment(StyleConstants.ALIGN_LEFT));
        btnAlignCenter = makeIconButton(" ", "Centrar",           e -> applyAlignment(StyleConstants.ALIGN_CENTER));
        btnAlignRight  = makeIconButton(" ", "Alinear derecha",   e -> applyAlignment(StyleConstants.ALIGN_RIGHT));
        tb.add(btnAlignLeft);
        tb.add(btnAlignCenter);
        tb.add(btnAlignRight);
        
        tb.addSeparator();
        
        btnColor = makeColorSwatch(" ", "Color de texto", Color.RED, e -> pickFontColor());
        btnBgColor = makeColorSwatch(" ", "Color de fondo", Color.YELLOW, e -> pickBgColor());
        tb.add(btnColor);
        tb.add(btnBgColor);
        
        tb.addSeparator();
        
        JButton btnTable = makeIconButton(" ", "Insertar tabla", e -> showTableDialog());
        tb.add(btnTable);

        tb.addSeparator();
        
        JButton btnOpen = makeIconButton(" ", "Abrir DOCX", e -> actionOpen());
        JButton btnSave = makeIconButton(" ", "Guardar DOCX", e -> actionSave());
        tb.add(btnOpen);
        tb.add(btnSave);
        
        return tb;
    }

    private JButton makeStyleButton(String text, String tip, int fontStyle, ActionListener al) {
        JButton b = new JButton(text);
        b.setToolTipText(tip);
        b.setFocusable(false);
        b.setPreferredSize(new Dimension(34, 28));
        b.addActionListener(al);
        return b;
    }
    
    private JButton makeIconButton(String text, String tip, ActionListener al) {
        JButton b = new JButton(text);
        b.setToolTipText(tip);
        b.setFocusable(false);
        b.setPreferredSize(new Dimension(34, 28));
        b.addActionListener(al);
        return b;
    }
    
    private JButton makeColorSwatch(String label, String tip, Color c, ActionListener al) {
        JButton b = new JButton(label) {
            @Override protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                g.setColor(c);
                g.fillRect(2, getHeight()-6, getWidth()-4, 4);
            }
        };
        b.setToolTipText(tip);
        b.setFocusable(false);
        b.setPreferredSize(new Dimension(34, 28));
        b.addActionListener(al);
        return b;
    }
    
    private void stylizeButton(JButton b, Font f) {
        b.setFont(f);
    }
    
    private void syncToolbarToSelection() {
        int pos = textPane.getCaretPosition();
        if (pos > 0) pos--;
        AttributeSet as = doc.getCharacterElement(pos).getAttributes();
        String fname = StyleConstants.getFontFamily(as);
        int fsize = StyleConstants.getFontSize(as);
        boolean bold = StyleConstants.isBold(as);
        boolean italic = StyleConstants.isItalic(as);
        boolean under = StyleConstants.isUnderline(as);
        
        fontNameBox.setSelectedItem(fname);
        fontSizeBox.setSelectedItem(fsize);
        btnBold.setBackground(bold ? new Color(180, 200, 255) : null);
        btnItalic.setBackground(italic ? new Color(180, 200, 255) : null);
        btnUnderline.setBackground(under ? new Color(180, 200, 255) : null);
    }
    
    private void applyFontName() {
        String name = (String) fontNameBox.getSelectedItem();
        if (name == null) return;
        applyStyle(StyleConstants.FontFamily, name);
    }
    
    private void applyFontSize() {
        Object sel = fontSizeBox.getSelectedItem();
        if (sel == null) return;
        try {
            int size = Integer.parseInt(sel.toString().trim());
            applyStyle(StyleConstants.FontSize, size);
        } catch (NumberFormatException ignored) {}
    }
    
    private void toggleBold() {
        int pos = textPane.getCaretPosition();
        if (pos > 0) pos--;
        boolean current = StyleConstants.isBold(doc.getCharacterElement(pos).getAttributes());
        applyStyle(StyleConstants.Bold, !current);
    }
    
    private void toggleItalic() {
        int pos = textPane.getCaretPosition();
        if (pos > 0) pos--;
        boolean current = StyleConstants.isItalic(doc.getCharacterElement(pos).getAttributes());
        applyStyle(StyleConstants.Italic, !current);
    }
    
    private void toggleUnderline() {
        int pos = textPane.getCaretPosition();
        if (pos > 0) pos--;
        boolean current = StyleConstants.isUnderline(doc.getCharacterElement(pos).getAttributes());
        applyStyle(StyleConstants.Underline, !current);
    }
    
    private void applyAlignment(int align) {
        MutableAttributeSet as = new SimpleAttributeSet();
        StyleConstants.setAlignment(as, align);
        int start = textPane.getSelectionStart();
        int end   = textPane.getSelectionEnd();
        doc.setParagraphAttributes(start, end - start, as, false);
    }
    
    private void pickFontColor() {
        Color c = JColorChooser.showDialog(this, "Color de texto", currentFontColor);
        if (c != null) {
            currentFontColor = c;
            applyStyle(StyleConstants.Foreground, c);
        }
    }
    
    private void pickBgColor() {
        Color c = JColorChooser.showDialog(this, "Color de resaltado", Color.YELLOW);
        if (c != null) {
            currentBgColor = c;
            applyStyle(StyleConstants.Background, c);
        }
    }
    
    @SuppressWarnings("unchecked")
    private <T> void applyStyle(Object key, T value) {
        int start = textPane.getSelectionStart();
        int end   = textPane.getSelectionEnd();
        MutableAttributeSet as = new SimpleAttributeSet();
        if      (key == StyleConstants.FontFamily) StyleConstants.setFontFamily(as, (String) value);
        else if (key == StyleConstants.FontSize)   StyleConstants.setFontSize(as, (Integer) value);
        else if (key == StyleConstants.Bold)       StyleConstants.setBold(as, (Boolean) value);
        else if (key == StyleConstants.Italic)     StyleConstants.setItalic(as, (Boolean) value);
        else if (key == StyleConstants.Underline)  StyleConstants.setUnderline(as, (Boolean) value);
        else if (key == StyleConstants.Foreground) StyleConstants.setForeground(as, (Color) value);
        else if (key == StyleConstants.Background) StyleConstants.setBackground(as, (Color) value);

        if (start == end) {
            
            textPane.setCharacterAttributes(as, false);
        } else {
            
            doc.setCharacterAttributes(start, end - start, as, false);
        }
        textPane.requestFocus();
    }
    
    //
    private void showFontDialog() {
        FontDialog fd = new FontDialog(this, textPane);
        fd.setVisible(true);
    }

    private void showTableDialog() {
        TableInsertDialog td = new TableInsertDialog(this, textPane, doc);
        td.setVisible(true);
    }

    private void actionNew() {
        if (confirmDiscard()) {
            textPane.setText("");
            currentFile = null;
            setTitle("Editor de Texto DOCX — Nuevo");
            statusBar.setText("  Nuevo documento");
        }
    }

    private void actionOpen() {
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Documentos Word (*.docx)", "docx"));
        fc.addChoosableFileFilter(new FileNameExtensionFilter("Archivos RTF (*.rtf)", "rtf"));
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File f = fc.getSelectedFile();
        try {
            if (f.getName().toLowerCase().endsWith(".docx")) {
                DocxHandler.load(f, textPane);
            } else {
                
                try (FileInputStream fis = new FileInputStream(f)) {
                    new RTFEditorKit().read(fis, doc, 0);
                }
            }
            
            currentFile = f;
            setTitle("Editor de Texto DOCX  " + f.getName());
            statusBar.setText("  Archivo abierto: " + f.getAbsolutePath());
            
        } catch (Exception ex) {
            
            JOptionPane.showMessageDialog(this, "Error al abrir: " + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    
    private void actionSave() {
        if (currentFile == null) { actionSaveAs(); return; }
        saveToFile(currentFile);
    }

    private void actionSaveAs() {
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Documentos Word (*.docx)", "docx"));
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File f = fc.getSelectedFile();
        if (!f.getName().toLowerCase().endsWith(".docx"))
            f = new File(f.getAbsolutePath() + ".docx");
        currentFile = f;
        saveToFile(f);
    }

    private void saveToFile(File f) {
        try {
            DocxHandler.save(f, textPane);
            setTitle("Editor de Texto DOCX  " + f.getName());
            statusBar.setText("  Guardado: " + f.getAbsolutePath());
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error al guardar: " + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private boolean confirmDiscard() {
        if (doc.getLength() == 0) return true;
        return JOptionPane.showConfirmDialog(this,
                "¿Descartar cambios actuales?", "Confirmación",
                JOptionPane.YES_NO_OPTION) == JOptionPane.YES_OPTION;
    }
    //
    
    public static void main(String[] args) {
        try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); }
        catch (Exception ignored) {}
        SwingUtilities.invokeLater(GUI::new);
    }
    
    static class FontComboRenderer extends DefaultListCellRenderer {
        @Override
        public Component getListCellRendererComponent(JList<?> list, Object value,
                int index, boolean isSelected, boolean cellHasFocus) {
            super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
            if (value != null) {
                String name = value.toString();
                try { setFont(new Font(name, Font.PLAIN, 14)); }
                catch (Exception ignored) {}
                setText(name);
            }
            return this;
        }
    }
}