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
import ru.citlab24.protokol.protocolmap.area.AreaPrimaryPanel;
import ru.citlab24.protokol.protocolmap.area.AreaProtocolPanel;
import ru.citlab24.protokol.protocolmap.house.ProtocolMapPanel;
import ru.citlab24.protokol.tabs.resourceTab.EquipmentTab;
import ru.citlab24.protokol.tabs.resourceTab.CalendarTab;
import ru.citlab24.protokol.tabs.resourceTab.PersonnelTab;
import ru.citlab24.protokol.tabs.qms.ShewhartMapTab;
import ru.citlab24.protokol.tabs.qms.VlkTab;


import com.formdev.flatlaf.FlatClientProperties;

import javax.swing.*;
import java.awt.*;

public class MainFrame extends JFrame {

    private final Building building = new Building();
    private final JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
    private final CardLayout cardLayout = new CardLayout();
    private final JPanel cardPanel = new JPanel(cardLayout);
    private BuildingTab buildingTab;
    private AreaProtocolPanel areaProtocolPanel;
    private String currentCard;

    private static final String CARD_PROTOCOL_HOME = "protocol-home";
    private static final String CARD_PROTOCOL_AREA = "protocol-area";
    private static final String CARD_PROTOCOL_MAP = "protocol-map";
    private static final String CARD_PROTOCOL_REQUEST = "protocol-request";
    private static final String CARD_RESOURCE_PERSONNEL = "resource-personnel";
    private static final String CARD_RESOURCE_EQUIPMENT = "resource-equipment";
    private static final String CARD_RESOURCE_CALENDAR = "resource-calendar";
    private static final String CARD_QMS_AUDIT = "qms-audit";
    private static final String CARD_QMS_NONCONFORMITIES = "qms-nonconformities";
    private static final String CARD_QMS_VLK = "qms-vlk";
    private static final String CARD_QMS_SHEWHART_MAP = "qms-shewhart-map";


    public MainFrame() {
        super();
        setProjectTitle(building.getName());
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setMinimumSize(new Dimension(1100, 720));
        setLocationByPlatform(true);

        // Современные оконные декорации (если FlatLaf активен)
        getRootPane().putClientProperty("JRootPane.titleBarBackground", AppTheme.MENU);
        getRootPane().putClientProperty("JRootPane.titleBarForeground", Color.WHITE);

        configureTabbedPane();
        initUI();

        pack();
        setLocationRelativeTo(null);
        setExtendedState(getExtendedState() | JFrame.MAXIMIZED_BOTH);
    }

    private JMenuBar createMenuBar(Runnable onLoadProject, Runnable onSaveProject) {
        JMenuBar menuBar = new JMenuBar();
        menuBar.setOpaque(true);
        menuBar.setBackground(AppTheme.MENU);
        menuBar.setBorder(BorderFactory.createEmptyBorder(6, 12, 6, 12));

        menuBar.add(styleTopMenu(createFileMenu()));
        menuBar.add(Box.createHorizontalStrut(8));
        menuBar.add(styleTopMenu(createProjectMenu(onLoadProject, onSaveProject)));
        menuBar.add(Box.createHorizontalStrut(8));
        menuBar.add(styleTopMenu(createResourceMenu()));
        menuBar.add(Box.createHorizontalStrut(8));
        menuBar.add(styleTopMenu(createQmsMenu()));
        menuBar.add(Box.createHorizontalStrut(12));
        JLabel versionLabel = new JLabel("v 1.4.1");
        versionLabel.setForeground(new Color(0xC6D2DC));
        versionLabel.setBorder(BorderFactory.createEmptyBorder(0, 4, 0, 0));
        menuBar.add(versionLabel);
        return menuBar;
    }

    private JMenu styleTopMenu(JMenu menu) {
        menu.setOpaque(false);
        menu.setForeground(Color.WHITE);
        menu.setBorder(BorderFactory.createEmptyBorder(4, 10, 4, 10));
        return menu;
    }

    private void configureTabbedPane() {
        // «Воздушные» вкладки — только свойства, гарантированно присутствующие во FlatLaf
        tabbedPane.putClientProperty(FlatClientProperties.TABBED_PANE_TAB_HEIGHT, 28);
        tabbedPane.putClientProperty(FlatClientProperties.TABBED_PANE_HAS_FULL_BORDER, Boolean.FALSE);
        tabbedPane.putClientProperty(FlatClientProperties.STYLE,
                "tabType: card; showTabSeparators: false; tabArc: 8; contentSeparatorHeight: 0");

        // Скролл, если вкладок много
        tabbedPane.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
    }

    private void initUI() {
        buildingTab = new BuildingTab(building);
        areaProtocolPanel = new AreaProtocolPanel();
        setJMenuBar(createMenuBar(this::requestLoadProject, this::requestSaveProject));

        tabbedPane.addTab("Титульная страница",    new TitlePageTab(building));
        tabbedPane.addTab("Характеристики здания", buildingTab);
        tabbedPane.addTab("Микроклимат",           new MicroclimateTab());
        tabbedPane.addTab("Вентиляция",            new VentilationTab(building));
        tabbedPane.addTab("Ионизирующее излучение",new RadiationTab());
        tabbedPane.addTab("КЕО",                   new LightingTab(building));
        tabbedPane.addTab("Освещение",             new ArtificialLightingTab(building));
        tabbedPane.addTab("Осв улица",             new StreetLightingTab(building));
        tabbedPane.addTab("Шумы",                  new NoiseTab(building));
        AppTheme.decorateWorkspace(tabbedPane);

        cardPanel.add(createScenePanel(tabbedPane), CARD_PROTOCOL_HOME);
        AppTheme.decorateWorkspace(areaProtocolPanel);
        cardPanel.add(createScenePanel(areaProtocolPanel), CARD_PROTOCOL_AREA);
        ProtocolMapPanel protocolMapPanel = new ProtocolMapPanel();
        AppTheme.decorateWorkspace(protocolMapPanel);
        cardPanel.add(createScenePanel(protocolMapPanel), CARD_PROTOCOL_MAP);
        AreaPrimaryPanel areaPrimaryPanel = new AreaPrimaryPanel();
        AppTheme.decorateWorkspace(areaPrimaryPanel);
        cardPanel.add(createScenePanel(areaPrimaryPanel), CARD_PROTOCOL_REQUEST);
        CalendarTab calendarTab = new CalendarTab();
        VlkTab vlkTab = new VlkTab(calendarTab::refreshEvents);

        PersonnelTab personnelTab = new PersonnelTab();
        EquipmentTab equipmentTab = new EquipmentTab();
        AppTheme.decorateWorkspace(personnelTab);
        AppTheme.decorateWorkspace(equipmentTab);
        AppTheme.decorateWorkspace(calendarTab);
        AppTheme.decorateWorkspace(vlkTab);
        cardPanel.add(createScenePanel(personnelTab), CARD_RESOURCE_PERSONNEL);
        cardPanel.add(createScenePanel(equipmentTab), CARD_RESOURCE_EQUIPMENT);
        cardPanel.add(createScenePanel(calendarTab), CARD_RESOURCE_CALENDAR);
        cardPanel.add(createPlaceholderScene(), CARD_QMS_AUDIT);
        cardPanel.add(createPlaceholderScene(), CARD_QMS_NONCONFORMITIES);
        cardPanel.add(createScenePanel(vlkTab), CARD_QMS_VLK);
        ShewhartMapTab shewhartMapTab = new ShewhartMapTab();
        AppTheme.decorateWorkspace(shewhartMapTab);
        cardPanel.add(createScenePanel(shewhartMapTab), CARD_QMS_SHEWHART_MAP);

        getContentPane().setLayout(new BorderLayout());
        getContentPane().add(cardPanel, BorderLayout.CENTER);
        showCard(CARD_PROTOCOL_HOME);
    }

    private void showCard(String cardName) {
        currentCard = cardName;
        cardLayout.show(cardPanel, cardName);
    }

    private void requestLoadProject() {
        if (CARD_PROTOCOL_AREA.equals(currentCard) && areaProtocolPanel != null) {
            areaProtocolPanel.requestLoadProject();
            String projectName = areaProtocolPanel.getProjectName();
            if (!projectName.isBlank()) {
                setProjectTitle(projectName);
            }
            return;
        }
        if (buildingTab != null) {
            buildingTab.requestLoadProject();
        }
    }

    private void requestSaveProject() {
        if (CARD_PROTOCOL_AREA.equals(currentCard) && areaProtocolPanel != null) {
            areaProtocolPanel.requestSaveProject();
            String projectName = areaProtocolPanel.getProjectName();
            if (!projectName.isBlank()) {
                setProjectTitle(projectName);
            }
            return;
        }
        if (buildingTab != null) {
            buildingTab.requestSaveProject();
        }
    }

    private JPanel createScenePanel(JComponent content) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(AppTheme.PANEL);
        panel.setBorder(BorderFactory.createEmptyBorder(12, 12, 12, 12));
        content.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(AppTheme.BORDER),
                BorderFactory.createEmptyBorder(6, 6, 6, 6)
        ));
        panel.add(content, BorderLayout.CENTER);
        return panel;
    }

    private JPanel createPlaceholderScene() {
        JLabel label = new JLabel("В разработке", SwingConstants.CENTER);
        label.setFont(label.getFont().deriveFont(Font.BOLD, 22f));
        JPanel panel = new JPanel(new BorderLayout());
        panel.add(label, BorderLayout.CENTER);
        return createScenePanel(panel);
    }

    private JMenu createFileMenu() {
        JMenu fileMenu = new JMenu("Файл");

        JMenuItem protocolHome = new JMenuItem("Заполнение протокола дом");
        protocolHome.addActionListener(e -> showCard(CARD_PROTOCOL_HOME));

        JMenuItem protocolArea = new JMenuItem("Заполнение протокола участок");
        protocolArea.addActionListener(e -> showCard(CARD_PROTOCOL_AREA));

        JMenuItem protocolMap = new JMenuItem("Сформировать первичку по домам");
        protocolMap.addActionListener(e -> showCard(CARD_PROTOCOL_MAP));

        JMenuItem protocolRequest = new JMenuItem("Сформировать первичку по участкам");
        protocolRequest.addActionListener(e -> showCard(CARD_PROTOCOL_REQUEST));

        fileMenu.add(protocolHome);
        fileMenu.add(protocolArea);
        fileMenu.add(protocolMap);
        fileMenu.add(protocolRequest);

        return fileMenu;
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


    private JMenu createResourceMenu() {
        JMenu resourceMenu = new JMenu("Ресурс");

        JMenuItem personnelItem = new JMenuItem("Персонал");
        personnelItem.addActionListener(e -> showCard(CARD_RESOURCE_PERSONNEL));

        JMenuItem equipmentItem = new JMenuItem("Оборудование");
        equipmentItem.addActionListener(e -> showCard(CARD_RESOURCE_EQUIPMENT));

        JMenuItem calendarItem = new JMenuItem("Календарь");
        calendarItem.addActionListener(e -> showCard(CARD_RESOURCE_CALENDAR));

        resourceMenu.add(personnelItem);
        resourceMenu.add(equipmentItem);
        resourceMenu.add(calendarItem);
        return resourceMenu;
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

    private JMenu createQmsMenu() {
        JMenu qmsMenu = new JMenu("СМК");

        JMenuItem auditItem = new JMenuItem("Аудит");
        auditItem.addActionListener(e -> showCard(CARD_QMS_AUDIT));

        JMenuItem nonconformitiesItem = new JMenuItem("Несоответствия");
        nonconformitiesItem.addActionListener(e -> showCard(CARD_QMS_NONCONFORMITIES));

        JMenuItem vlkItem = new JMenuItem("ВЛК");
        vlkItem.addActionListener(e -> showCard(CARD_QMS_VLK));

        JMenuItem shewhartMapItem = new JMenuItem("Карта Шухарта");
        shewhartMapItem.addActionListener(e -> showCard(CARD_QMS_SHEWHART_MAP));

        qmsMenu.add(auditItem);
        qmsMenu.add(nonconformitiesItem);
        qmsMenu.add(vlkItem);
        qmsMenu.add(shewhartMapItem);
        return qmsMenu;
    }
}
