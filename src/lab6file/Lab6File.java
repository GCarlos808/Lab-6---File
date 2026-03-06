package lab6file;

import javax.swing.SwingUtilities;
import javax.swing.UIManager;

public class Lab6File {

    public static void main(String[] args) {
        try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); }
        catch (Exception ignored) {}
        SwingUtilities.invokeLater(GUI::new);
    }
}