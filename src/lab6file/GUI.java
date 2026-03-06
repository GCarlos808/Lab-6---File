package lab6file;

import javax.swing.*;
import javax.swing.text.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import javax.swing.undo.UndoManager;



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
    private Color currentBgColor   = Color.YELLOW;
    private final UndoManager undoManager = new UndoManager();
    
    private TextoFormat textoFormat;
    private GestorTablas gestorTablas;
    
    private static final String[] FONTS = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
    
    private static final Integer[] SIZES = {8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40, 48, 56, 64, 72, 96, 144, 190, 240, 300};
    
    public GUI() {
        
        super("Editor de Texto DOCX");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(1100, 750);
        setLocationRelativeTo(null);
        
        textPane = new JTextPane();
        textPane.setFont(new Font("Arial", Font.PLAIN, 12));
        doc = textPane.getStyledDocument();
        
        textoFormat  = new TextoFormat(textPane);
        gestorTablas = new GestorTablas(textPane);
        doc.addUndoableEditListener(e -> undoManager.addEdit(e.getEdit()));
        
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
        addItem(mFormat, "Negrita",           0, e -> textoFormat.negrita());
        addItem(mFormat, "Cursiva",           0, e -> textoFormat.cursiva());
        addItem(mFormat, "Subrayado",         0, e -> textoFormat.subrayado());
        addItem(mFormat, "Color de texto...", 0, e -> pickFontColor());
        addItem(mFormat, "Color de fondo...", 0, e -> pickBgColor());
        mb.add(mFormat);
        
        JMenu mTable = new JMenu("Tabla");
        addItem(mTable, "Insertar tabla...", 0, e -> gestorTablas.insertarTabla());
        mb.add(mTable);
        
        JMenu mHelp = new JMenu("Ayuda");
        addItem(mHelp, "Acerca de...", 0, e ->
                JOptionPane.showMessageDialog(this,
                        "Editor de Texto DOCX\nSoporta: fuentes, colores, tamaños, tablas\nFormato .docx con Apache POI",
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
        
        btnBold      = makeStyleButton("B", "Negrita (Ctrl+B)",   Font.BOLD,   e -> textoFormat.negrita());
        btnItalic    = makeStyleButton("I", "Cursiva (Ctrl+I)",   Font.ITALIC, e -> textoFormat.cursiva());
        btnUnderline = makeStyleButton("U", "Subrayado (Ctrl+U)", Font.PLAIN,  e -> textoFormat.subrayado());
        
        stylizeButton(btnBold,      new Font("Serif", Font.BOLD,   14));
        stylizeButton(btnItalic,    new Font("Serif", Font.ITALIC, 14));
        stylizeButton(btnUnderline, new Font("Serif", Font.PLAIN,  14));
        
        tb.add(btnBold);
        tb.add(btnItalic);
        tb.add(btnUnderline);
        
        bindKey("control B", "bold",      e -> textoFormat.negrita());
        bindKey("control I", "italic",    e -> textoFormat.cursiva());
        bindKey("control U", "underline", e -> textoFormat.subrayado());
        
        tb.addSeparator();
        
        btnAlignLeft   = makeIconButton("≡←", "Alinear izquierda", e -> applyAlignment(StyleConstants.ALIGN_LEFT));
        btnAlignCenter = makeIconButton("≡≡",  "Centrar",           e -> applyAlignment(StyleConstants.ALIGN_CENTER));
        btnAlignRight  = makeIconButton("→≡", "Alinear derecha",   e -> applyAlignment(StyleConstants.ALIGN_RIGHT));
        tb.add(btnAlignLeft);
        tb.add(btnAlignCenter);
        tb.add(btnAlignRight);
        
        tb.addSeparator();
        
        btnColor   = makeColorSwatchButton("A",  "Color de texto",   currentFontColor, e -> pickFontColor());
        btnBgColor = makeColorSwatchButton("AB", "Color de fondo",   currentBgColor,   e -> pickBgColor());
        tb.add(btnColor);
        tb.add(btnBgColor);
        
        tb.addSeparator();
        
        JButton btnTable = makeIconButton("#", "Insertar tabla", e -> gestorTablas.insertarTabla());
        tb.add(btnTable);
        
        tb.addSeparator();
        
        JButton btnOpen = makeIconButton("A", "Abrir DOCX", e -> actionOpen());
        JButton btnSave = makeIconButton("S", "Guardar DOCX", e -> actionSave());
        tb.add(btnOpen);
        tb.add(btnSave);

        tb.addSeparator();
        
        tb.addSeparator();
        
        JButton btnCut   = makeIconButton("X", "Cortar (Ctrl+X)",   e -> textPane.cut());
        JButton btnCopy  = makeIconButton("C", "Copiar (Ctrl+C)",   e -> textPane.copy());
        JButton btnPaste = makeIconButton("V", "Pegar (Ctrl+V)",   e -> textPane.paste());
        tb.add(btnCut);
        tb.add(btnCopy);
        tb.add(btnPaste);
        
        tb.addSeparator();
        
        JButton btnUndo = makeIconButton("↩", "Deshacer (Ctrl+Z)", e -> { if (undoManager.canUndo()) undoManager.undo(); });
        JButton btnRedo = makeIconButton("↪", "Rehacer (Ctrl+Y)", e -> { if (undoManager.canRedo()) undoManager.redo(); });
        tb.add(btnUndo);
        tb.add(btnRedo);
        
        bindKey("control Z", "Deshacer", e -> { if (undoManager.canUndo()) undoManager.undo(); });
        bindKey("control Y", "Rehacer", e -> { if (undoManager.canRedo()) undoManager.redo(); });
        return tb;
        
    }
    
    private void bindKey(String keyStroke, String actionName, ActionListener al) {
        textPane.getInputMap().put(KeyStroke.getKeyStroke(keyStroke), actionName);
        textPane.getActionMap().put(actionName, new AbstractAction() {
            @Override public void actionPerformed(ActionEvent e) { al.actionPerformed(e); }
        });
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
        b.setPreferredSize(new Dimension(40, 28));
        b.addActionListener(al);
        return b;
    }
    
    private JButton makeColorSwatchButton(String label, String tip, Color initialColor, ActionListener al) {
        JButton b = new JButton(label) {
            
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                Color c = (Color) getClientProperty("swatchColor");
                if (c == null) c = Color.BLACK;
                g.setColor(c);
                g.fillRect(3, getHeight() - 5, getWidth() - 6, 4);
            }
        };
        
        b.putClientProperty("swatchColor", initialColor);
        b.setToolTipText(tip);
        b.setFocusable(false);
        b.setPreferredSize(new Dimension(40, 28));
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
        int    fsize = StyleConstants.getFontSize(as);
        boolean bold   = StyleConstants.isBold(as);
        boolean italic = StyleConstants.isItalic(as);
        boolean under  = StyleConstants.isUnderline(as);
        
        fontNameBox.setSelectedItem(fname);
        fontSizeBox.setSelectedItem(fsize);
        btnBold.setBackground(bold   ? new Color(180, 200, 255) : null);
        btnItalic.setBackground(italic ? new Color(180, 200, 255) : null);
        btnUnderline.setBackground(under  ? new Color(180, 200, 255) : null);
    }
    
    private void applyFontName() {
        String name = (String) fontNameBox.getSelectedItem();
        if (name != null) textoFormat.fuente(name);
    }
    
    private void applyFontSize() {
        Object sel = fontSizeBox.getSelectedItem();
        if (sel != null) textoFormat.tamano(sel.toString());
    }
    
    private void pickFontColor() {
        Color c = JColorChooser.showDialog(this, "Color de texto", currentFontColor);
        if (c != null) {
            currentFontColor = c;
            btnColor.putClientProperty("swatchColor", c);
            btnColor.repaint();
            applyStyle(StyleConstants.Foreground, c);
        }
    }
    
    private void pickBgColor() {
        Color c = JColorChooser.showDialog(this, "Color de resaltado", currentBgColor);
        if (c != null) {
            currentBgColor = c;
            btnBgColor.putClientProperty("swatchColor", c);
            btnBgColor.repaint();
            applyStyle(StyleConstants.Background, c);
        }
    }
    
    private void applyAlignment(int align) {
        MutableAttributeSet as = new SimpleAttributeSet();
        StyleConstants.setAlignment(as, align);
        int start = textPane.getSelectionStart();
        int end   = textPane.getSelectionEnd();
        doc.setParagraphAttributes(start, end - start, as, false);
    }
    
    @SuppressWarnings("unchecked")
    private <T> void applyStyle(Object key, T value) {
        int start = textPane.getSelectionStart();
        int end   = textPane.getSelectionEnd();
        MutableAttributeSet as = new SimpleAttributeSet();
        if      (key == StyleConstants.Foreground) StyleConstants.setForeground(as, (Color) value);
        else if (key == StyleConstants.Background) StyleConstants.setBackground(as, (Color) value);

        if (start == end) {
            textPane.setCharacterAttributes(as, false);
        } else {
            doc.setCharacterAttributes(start, end - start, as, false);
        }
        textPane.requestFocus();
    }
    
    private void actionNew() {
        if (confirmDiscard()) {
            try {
                doc.remove(0, doc.getLength());
            } catch (BadLocationException ignored) {}
            gestorTablas.limpiar();
            currentFile = null;
            setTitle("Editor de Texto DOCX — Nuevo");
            statusBar.setText("  Nuevo documento");
        }
    }
    
    private void actionOpen() {
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Documentos Word (*.docx)", "docx"));
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION)
            return;
        
        File f = fc.getSelectedFile();
        gestorTablas.limpiar();
        GestorArchivos.abrir(textPane, f.getAbsolutePath());
        currentFile = f;
        setTitle("Editor de Texto DOCX — " + f.getName());
        statusBar.setText("  Archivo abierto: " + f.getAbsolutePath());
    }
    
    private void actionSave() {
        if (currentFile == null) {
            actionSaveAs();
            return;
        }
        saveToFile(currentFile);
    }
    
    private void actionSaveAs() {
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Documentos Word (*.docx)", "docx"));
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION)
            return;
        
        File f = fc.getSelectedFile();
        if (!f.getName().toLowerCase().endsWith(".docx"))
            f = new File(f.getAbsolutePath() + ".docx");
        
        currentFile = f;
        saveToFile(f);
    }
    
    private void saveToFile(File f) {
        GestorArchivos.guardar(textPane, f.getAbsolutePath());
        setTitle("Editor de Texto DOCX — " + f.getName());
        statusBar.setText("  Guardado: " + f.getAbsolutePath());
    }
    
    private boolean confirmDiscard() {
        if (doc.getLength() == 0) return true;
        return JOptionPane.showConfirmDialog(this,
                "¿Descartar cambios actuales?", "Confirmación",JOptionPane.YES_NO_OPTION) == JOptionPane.YES_OPTION;
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