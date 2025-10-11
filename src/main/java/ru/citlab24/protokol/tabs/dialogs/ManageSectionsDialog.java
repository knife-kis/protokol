package ru.citlab24.protokol.tabs.dialogs;

import ru.citlab24.protokol.tabs.models.Section;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;
import java.util.List;

public class ManageSectionsDialog extends JDialog {
    private boolean confirmed = false;
    private final JPanel namesPanel = new JPanel(new GridLayout(0, 1, 6, 6));
    private final List<JTextField> nameFields = new ArrayList<>();
    private final JSpinner countSpinner;

    public ManageSectionsDialog(JFrame parent, List<Section> current) {
        super(parent, "Секции (подъезды)", true);
        setSize(380, 360);
        setLocationRelativeTo(parent);
        setLayout(new BorderLayout(10,10));

        int count = Math.max(1, current.size());
        countSpinner = new JSpinner(new SpinnerNumberModel(count, 1, 50, 1));
        countSpinner.addChangeListener(e -> rebuildNames((Integer) countSpinner.getValue(), null));

        JPanel top = new JPanel(new FlowLayout(FlowLayout.LEFT));
        top.add(new JLabel("Количество секций:"));
        top.add(countSpinner);
        add(top, BorderLayout.NORTH);

        add(new JScrollPane(namesPanel), BorderLayout.CENTER);

        JPanel btns = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton ok = new JButton("ОК");
        JButton cancel = new JButton("Отмена");
        btns.add(ok); btns.add(cancel);
        add(btns, BorderLayout.SOUTH);

        ok.addActionListener(e -> { confirmed = true; setVisible(false); });
        cancel.addActionListener(e -> { confirmed = false; setVisible(false); });

        // начальная раскладка имён
        rebuildNames(count, current);
    }

    private void rebuildNames(int count, List<Section> preset) {
        namesPanel.removeAll();
        nameFields.clear();
        for (int i = 0; i < count; i++) {
            String val = (preset != null && i < preset.size())
                    ? preset.get(i).getName()
                    : "Секция " + (i + 1);
            JTextField tf = new JTextField(val);
            nameFields.add(tf);
            JPanel row = new JPanel(new BorderLayout(6,0));
            row.add(new JLabel((i+1) + ":"), BorderLayout.WEST);
            row.add(tf, BorderLayout.CENTER);
            namesPanel.add(row);
        }
        namesPanel.revalidate();
        namesPanel.repaint();
    }

    public boolean showDialog() { setVisible(true); return confirmed; }

    public List<Section> getSections() {
        List<Section> res = new ArrayList<>();
        for (int i = 0; i < nameFields.size(); i++) {
            Section s = new Section(nameFields.get(i).getText().trim(), i);
            res.add(s);
        }
        return res;
    }
}
