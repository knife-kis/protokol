package ru.citlab24.protokol;

import ru.citlab24.protokol.tabs.buildingTab.BuildingTab;
import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.modules.lighting.ArtificialLightingTab;
import ru.citlab24.protokol.tabs.modules.lighting.LightingTab;
import ru.citlab24.protokol.tabs.modules.lighting.StreetLightingTab;
import ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateTab;
import ru.citlab24.protokol.tabs.modules.noise.NoiseTab;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationTab;
import ru.citlab24.protokol.tabs.modules.med.RadiationTab;
import ru.citlab24.protokol.tabs.titleTab.TitlePageTab;


import com.formdev.flatlaf.FlatLaf;
import com.formdev.flatlaf.FlatDarkLaf;
import com.formdev.flatlaf.FlatLightLaf;
import com.formdev.flatlaf.FlatDarculaLaf;
import com.formdev.flatlaf.FlatIntelliJLaf;
import com.formdev.flatlaf.FlatClientProperties;

import javax.swing.*;
import java.awt.*;

public class MainFrame extends JFrame {

    private final Building building = new Building();
    private final JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);


    public MainFrame() {
        super("Протокол испытаний");
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setMinimumSize(new Dimension(1100, 720));
        setLocationByPlatform(true);

        // Современные оконные декорации (если FlatLaf активен)
        getRootPane().putClientProperty("JRootPane.titleBarBackground", UIManager.getColor("Panel.background"));
        getRootPane().putClientProperty("JRootPane.titleBarForeground", UIManager.getColor("Label.foreground"));

        setJMenuBar(createMenuBar());
        configureTabbedPane();
        initUI();

        pack();
        setLocationRelativeTo(null);
    }

    private JMenuBar createMenuBar() {
        JMenuBar menuBar = new JMenuBar();

        // ===== Меню «Вид» -> «Тема» =====
        JMenu viewMenu = new JMenu("Вид");
        JMenu themeMenu = new JMenu("Тема");

        JMenuItem light = new JMenuItem("Светлая");
        JMenuItem dark = new JMenuItem("Тёмная");
        JMenuItem intellij = new JMenuItem("IntelliJ");
        JMenuItem darcula = new JMenuItem("Darcula");
        JMenuItem system = new JMenuItem("Системная (классическая)");

        light.addActionListener(e -> switchLaf(new FlatLightLaf()));
        dark.addActionListener(e -> switchLaf(new FlatDarkLaf()));
        intellij.addActionListener(e -> switchLaf(new FlatIntelliJLaf()));
        darcula.addActionListener(e -> switchLaf(new FlatDarculaLaf()));
        system.addActionListener(e -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                FlatLaf.updateUI();
                SwingUtilities.updateComponentTreeUI(this);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });

        themeMenu.add(light);
        themeMenu.add(dark);
        themeMenu.add(intellij);
        themeMenu.add(darcula);
        themeMenu.addSeparator();
        themeMenu.add(system);

        // (Опционально) настройки вкладок
        JMenu tabsMenu = new JMenu("Вкладки");
        JCheckBoxMenuItem compactTabs = new JCheckBoxMenuItem("Компактный режим");
        compactTabs.addActionListener(e -> {
            boolean compact = compactTabs.isSelected();
            tabbedPane.putClientProperty(FlatClientProperties.TABBED_PANE_TAB_HEIGHT, compact ? 28 : 40);
            tabbedPane.revalidate();
            tabbedPane.repaint();
        });

        tabsMenu.add(compactTabs);
        viewMenu.add(themeMenu);
        viewMenu.add(tabsMenu);

        menuBar.add(viewMenu);
        return menuBar;
    }

    private void switchLaf(LookAndFeel laf) {
        try {
            UIManager.setLookAndFeel(laf);
            FlatLaf.updateUI();
            SwingUtilities.updateComponentTreeUI(this);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void configureTabbedPane() {
        // «Воздушные» вкладки — только свойства, гарантированно присутствующие во FlatLaf
        tabbedPane.putClientProperty(FlatClientProperties.TABBED_PANE_TAB_HEIGHT, 40);
        tabbedPane.putClientProperty(FlatClientProperties.TABBED_PANE_HAS_FULL_BORDER, Boolean.FALSE);

        // Скролл, если вкладок много
        tabbedPane.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);

        getContentPane().setLayout(new BorderLayout());
        getContentPane().add(tabbedPane, BorderLayout.CENTER);
    }

    private void initUI() {
        tabbedPane.addTab("Титульная страница",    new TitlePageTab(building));
        tabbedPane.addTab("Характеристики здания", new BuildingTab(building));
        tabbedPane.addTab("Микроклимат",           new MicroclimateTab());
        tabbedPane.addTab("Вентиляция",            new VentilationTab(building));
        tabbedPane.addTab("Ионизирующее излучение",new RadiationTab());
        tabbedPane.addTab("КЕО",                   new LightingTab(building));
        tabbedPane.addTab("Освещение",             new ArtificialLightingTab(building));
        tabbedPane.addTab("Осв улица",             new StreetLightingTab(building));
        tabbedPane.addTab("Шумы",                  new NoiseTab(building));
    }

    // ===== Утилиты доступа к вкладкам =====
    public RadiationTab getRadiationTab() {
        for (Component comp : tabbedPane.getComponents()) {
            if (comp instanceof RadiationTab) return (RadiationTab) comp;
        }
        return null;
    }

    public BuildingTab getBuildingTab() {
        for (Component comp : tabbedPane.getComponents()) {
            if (comp instanceof BuildingTab) {
                return (BuildingTab) comp;
            }
        }
        return null;
    }

    // NEW: нужен AllExcelExporter
    public VentilationTab getVentilationTab() {
        for (Component comp : tabbedPane.getComponents()) {
            if (comp instanceof VentilationTab) return (VentilationTab) comp;
        }
        return null;
    }
    public ArtificialLightingTab getArtificialLightingTab() {
        for (Component comp : tabbedPane.getComponents()) {
            if (comp instanceof ArtificialLightingTab) return (ArtificialLightingTab) comp;
        }
        return null;
    }

    public void selectVentilationTab() {
        for (int i = 0; i < tabbedPane.getTabCount(); i++) {
            if ("Вентиляция".equals(tabbedPane.getTitleAt(i))) {
                tabbedPane.setSelectedIndex(i);
                return;
            }
        }
    }
    public StreetLightingTab getStreetLightingTab() {
        for (java.awt.Component c : getTabbedPane().getComponents()) {
            if (c instanceof StreetLightingTab) {
                return (StreetLightingTab) c;
            }
        }
        return null;
    }

    public NoiseTab getNoiseTab() {
        for (Component comp : tabbedPane.getComponents()) {
            if (comp instanceof NoiseTab) {
                return (NoiseTab) comp;
            }
        }
        return null;
    }

    public JTabbedPane getTabbedPane() {
        return tabbedPane;
    }
}
