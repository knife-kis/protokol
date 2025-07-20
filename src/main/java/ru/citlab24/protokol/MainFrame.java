package ru.citlab24.protokol;

import ru.citlab24.protokol.tabs.VentilationTab;
import ru.citlab24.protokol.tabs.buildingTab.BuildingTab;
import ru.citlab24.protokol.tabs.microclimateTab.MicroclimateTab;

import javax.swing.*;
import java.awt.*;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import ru.citlab24.protokol.tabs.models.Building;

public class MainFrame extends JFrame {
    private final Building building = new Building();
    private final JTabbedPane tabbedPane = new JTabbedPane();

    public MainFrame() {
        super("Building Analytics");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 750);
        setLocationRelativeTo(null);
        initUI();
        tabbedPane.addChangeListener(e -> {
            Component selectedTab = tabbedPane.getSelectedComponent();
            if (selectedTab instanceof VentilationTab) {
                ((VentilationTab) selectedTab).refreshData();
            } else if (selectedTab instanceof BuildingTab) {
                // Сохраняем изменения при переходе на вкладку здания
                for (Component tab : tabbedPane.getComponents()) {
                    if (tab instanceof VentilationTab) {
                        ((VentilationTab) tab).saveCalculationsToModel();
                    }
                }
            }
        });
    }

    private void initUI() {
        tabbedPane.addTab("Характеристики здания",
                FontIcon.of(FontAwesomeSolid.BUILDING, 24, new Color(0, 115, 200)),
                new BuildingTab(building)); // Обновлённый BuildingTab

        tabbedPane.addTab("Микроклимат",
                FontIcon.of(FontAwesomeSolid.THERMOMETER_HALF, 24, new Color(76, 175, 80)),
                new MicroclimateTab());

        tabbedPane.addTab("Вентиляция",
                FontIcon.of(FontAwesomeSolid.WIND, 24, new Color(41, 182, 246)), // Иконка вентиляции
                new VentilationTab(building)); // Прямо в главной панели

        add(tabbedPane, BorderLayout.CENTER);
    }
    public void selectVentilationTab() {
        tabbedPane.setSelectedIndex(2); // 2 - индекс вкладки "Вентиляция"
    }
    public JTabbedPane getTabbedPane() {
        return tabbedPane;
    }
}