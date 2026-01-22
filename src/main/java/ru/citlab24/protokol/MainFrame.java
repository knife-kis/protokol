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
    private final CardLayout cardLayout = new CardLayout();
    private final JPanel cardPanel = new JPanel(cardLayout);

    private static final String CARD_HOME = "home";
    private static final String CARD_PROTOCOL_HOME = "protocol-home";
    private static final String CARD_PROTOCOL_AREA = "protocol-area";
    private static final String CARD_PROTOCOL_MAP = "protocol-map";
    private static final String CARD_PROTOCOL_REQUEST = "protocol-request";


    public MainFrame() {
        super();
        setProjectTitle(building.getName());
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setMinimumSize(new Dimension(1100, 720));
        setLocationByPlatform(true);

        // Современные оконные декорации (если FlatLaf активен)
        getRootPane().putClientProperty("JRootPane.titleBarBackground", UIManager.getColor("Panel.background"));
        getRootPane().putClientProperty("JRootPane.titleBarForeground", UIManager.getColor("Label.foreground"));

        configureTabbedPane();
        initUI();

        pack();
        setLocationRelativeTo(null);
    }

    private JMenuBar createMenuBar(Runnable onLoadProject, Runnable onSaveProject) {
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
        compactTabs.setSelected(true);
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
        menuBar.add(Box.createHorizontalStrut(8));
        menuBar.add(createProjectMenu(onLoadProject, onSaveProject));
        menuBar.add(Box.createHorizontalStrut(12));
        JLabel versionLabel = new JLabel("v 1.1.2");
        versionLabel.setBorder(BorderFactory.createEmptyBorder(0, 4, 0, 0));
        menuBar.add(versionLabel);
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
        tabbedPane.putClientProperty(FlatClientProperties.TABBED_PANE_TAB_HEIGHT, 28);
        tabbedPane.putClientProperty(FlatClientProperties.TABBED_PANE_HAS_FULL_BORDER, Boolean.FALSE);

        // Скролл, если вкладок много
        tabbedPane.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
    }

    private void initUI() {
        BuildingTab buildingTab = new BuildingTab(building);
        setJMenuBar(createMenuBar(buildingTab::requestLoadProject, buildingTab::requestSaveProject));

        tabbedPane.addTab("Титульная страница",    new TitlePageTab(building));
        tabbedPane.addTab("Характеристики здания", buildingTab);
        tabbedPane.addTab("Микроклимат",           new MicroclimateTab());
        tabbedPane.addTab("Вентиляция",            new VentilationTab(building));
        tabbedPane.addTab("Ионизирующее излучение",new RadiationTab());
        tabbedPane.addTab("КЕО",                   new LightingTab(building));
        tabbedPane.addTab("Освещение",             new ArtificialLightingTab(building));
        tabbedPane.addTab("Осв улица",             new StreetLightingTab(building));
        tabbedPane.addTab("Шумы",                  new NoiseTab(building));

        cardPanel.add(createHomePanel(), CARD_HOME);
        cardPanel.add(createScenePanel("Заполнение протоколов — дом", tabbedPane), CARD_PROTOCOL_HOME);
        cardPanel.add(createPlaceholderScene("Заполнение протоколов — участок"), CARD_PROTOCOL_AREA);
        cardPanel.add(createPlaceholderScene("Сформировать карту по протоколу"), CARD_PROTOCOL_MAP);
        cardPanel.add(createPlaceholderScene("Сформировать заявку по протоколу"), CARD_PROTOCOL_REQUEST);

        getContentPane().setLayout(new BorderLayout());
        getContentPane().add(cardPanel, BorderLayout.CENTER);
        cardLayout.show(cardPanel, CARD_HOME);
    }

    private JPanel createHomePanel() {
        JPanel container = new JPanel(new BorderLayout());
        container.setBorder(BorderFactory.createEmptyBorder(32, 32, 32, 32));

        JPanel grid = new JPanel(new GridLayout(2, 2, 24, 24));
        grid.add(createHomeButton("Заполнение протоколов дом", CARD_PROTOCOL_HOME));
        grid.add(createHomeButton("Заполнение протоколов участок", CARD_PROTOCOL_AREA));
        grid.add(createHomeButton("Сформировать карту по протоколу", CARD_PROTOCOL_MAP));
        grid.add(createHomeButton("Сформировать заявку по протоколу", CARD_PROTOCOL_REQUEST));

        container.add(grid, BorderLayout.CENTER);
        return container;
    }

    private JButton createHomeButton(String title, String cardName) {
        JButton button = new JButton("<html><div style='text-align:center;'>" + title + "</div></html>");
        button.setFocusPainted(false);
        button.setFont(button.getFont().deriveFont(Font.BOLD, 18f));
        button.setPreferredSize(new Dimension(280, 180));
        button.addActionListener(e -> cardLayout.show(cardPanel, cardName));
        return button;
    }

    private JPanel createScenePanel(String title, JComponent content) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.add(createSceneHeader(title), BorderLayout.NORTH);
        panel.add(content, BorderLayout.CENTER);
        return panel;
    }

    private JPanel createPlaceholderScene(String title) {
        JLabel label = new JLabel("В разработке", SwingConstants.CENTER);
        label.setFont(label.getFont().deriveFont(Font.BOLD, 22f));
        JPanel panel = new JPanel(new BorderLayout());
        panel.add(label, BorderLayout.CENTER);
        return createScenePanel(title, panel);
    }

    private JPanel createSceneHeader(String title) {
        JPanel header = new JPanel(new BorderLayout());
        header.setBorder(BorderFactory.createEmptyBorder(12, 16, 12, 16));

        JButton backButton = new JButton("Назад");
        backButton.addActionListener(e -> cardLayout.show(cardPanel, CARD_HOME));

        JLabel titleLabel = new JLabel(title);
        titleLabel.setFont(titleLabel.getFont().deriveFont(Font.BOLD, 16f));

        header.add(backButton, BorderLayout.WEST);
        header.add(titleLabel, BorderLayout.CENTER);
        return header;
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

    public void setProjectTitle(String projectName) {
        String title = (projectName == null || projectName.isBlank())
                ? "новый проект"
                : projectName.trim();
        setTitle(title);
    }

    private JMenu createProjectMenu(Runnable onLoadProject, Runnable onSaveProject) {
        JMenu projectMenu = new JMenu("Проект");
        JMenuItem loadItem = new JMenuItem(
                "Загрузить проект",
                org.kordamp.ikonli.swing.FontIcon.of(
                        org.kordamp.ikonli.fontawesome5.FontAwesomeSolid.FOLDER_OPEN,
                        14
                )
        );
        JMenuItem saveItem = new JMenuItem(
                "Сохранить проект",
                org.kordamp.ikonli.swing.FontIcon.of(
                        org.kordamp.ikonli.fontawesome5.FontAwesomeSolid.SAVE,
                        14
                )
        );

        loadItem.addActionListener(event -> {
            Runnable action = (onLoadProject != null) ? onLoadProject : () -> {};
            action.run();
        });
        saveItem.addActionListener(event -> {
            Runnable action = (onSaveProject != null) ? onSaveProject : () -> {};
            action.run();
        });

        projectMenu.add(loadItem);
        projectMenu.add(saveItem);
        return projectMenu;
    }
}
