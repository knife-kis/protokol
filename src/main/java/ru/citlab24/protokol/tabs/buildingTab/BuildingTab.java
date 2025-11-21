package ru.citlab24.protokol.tabs.buildingTab;

import javafx.application.Platform;
import javafx.embed.swing.JFXPanel;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.db.ProjectFileService;
import ru.citlab24.protokol.export.AllExcelExporter;
import ru.citlab24.protokol.tabs.dialogs.*;
import ru.citlab24.protokol.tabs.modules.lighting.LightingTab;
import ru.citlab24.protokol.tabs.modules.lighting.ArtificialLightingTab;
import ru.citlab24.protokol.tabs.modules.lighting.StreetLightingTab;
import ru.citlab24.protokol.tabs.modules.med.RadiationTab;
import ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateTab;
import ru.citlab24.protokol.tabs.modules.noise.NoiseTab;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationTab;
import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;
import ru.citlab24.protokol.tabs.renderers.RoomListRenderer;
import ru.citlab24.protokol.tabs.renderers.SpaceListRenderer;


import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.awt.event.*;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class BuildingTab extends JPanel {
    private static final Logger logger = LoggerFactory.getLogger(BuildingTab.class);
    // Константы для цветов и стилей
    private static final java.awt.Color FLOOR_PANEL_COLOR = new java.awt.Color(0, 115, 200);
    private static final java.awt.Color SPACE_PANEL_COLOR = new java.awt.Color(76, 175, 80);
    private static final java.awt.Color ROOM_PANEL_COLOR  = new java.awt.Color(156, 39, 176);

    private static final Font HEADER_FONT =
            UIManager.getFont("Label.font").deriveFont(Font.PLAIN, 15f);
    private static final Dimension BUTTON_PANEL_SIZE = new Dimension(5, 5);

    private Building building;
    private BuildingModelOps ops;
    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final DefaultListModel<Space> spaceListModel = new DefaultListModel<>();
    private final DefaultListModel<Room> roomListModel = new DefaultListModel<>();
    private final DefaultListModel<Section> sectionListModel = new DefaultListModel<>();
    private JList<Section> sectionList;

    private JList<Floor> floorList;
    private JList<Space> spaceList;
    private JList<Room> roomList;
    private JTextField projectNameField;
    // === Фильтр типов помещений (Жилые / Офисные / Общественные) ===
    private JToggleButton filterApartmentBtn;
    private JToggleButton filterOfficeBtn;
    private JToggleButton filterPublicBtn;
    private JToggleButton filterStreetBtn; // НОВОЕ
    private boolean[] savedFilters = new boolean[]{true, true, true, true}; // 4 флага

    public BuildingTab(Building building) {
        // 1) сохраняем в поле, 2) создаём дефолт, если пришёл null
        this.building = (building != null) ? building : new Building();
        if (this.building.getSections().isEmpty()) {
            this.building.addSection(new Section("Секция 1", 0));
        }
        this.ops = new BuildingModelOps(this.building); // ← добавили
        initComponents();
    }

    private static String normalizeRoomBaseName(String name) {
        return BuildingModelOps.normalizeRoomBaseName(name);
    }

    private void initComponents() {
        setLayout(new BorderLayout());
        add(createProjectNamePanel(), BorderLayout.NORTH);
        add(createBuildingPanel(), BorderLayout.CENTER);
        add(createActionButtons(), BorderLayout.SOUTH);

        setupListSelectionListeners();
    }

    private void setupListSelectionListeners() {
        // Инициализация будет после создания компонентов
    }

    private JPanel createProjectNamePanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(0, 10, 0, 10));

        JLabel label = new JLabel("Название проекта:");
        label.setFont(HEADER_FONT);

        projectNameField = new JTextField(30);
        projectNameField.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        String bName = (this.building.getName() != null) ? this.building.getName() : "";
        projectNameField.setText(bName);

        panel.add(label);
        panel.add(projectNameField);

        return panel;
    }

    private JPanel createBuildingPanel() {
        JPanel panel = new JPanel(new GridLayout(1, 3, 15, 15));
        panel.add(createSectionsAndFloorsPanel());
        panel.add(createSpacesPanelWithFilter()); // ← ТУТ теперь панель с фильтром
        panel.add(createListPanel("Комнаты в помещении", ROOM_PANEL_COLOR, roomListModel,
                new RoomListRenderer(), this::createRoomButtons));
        return panel;
    }
    /** Средняя колонка: «Помещения на этаже» + панель фильтров (Жилые/Офисные/Общественные). */
    private JPanel createSpacesPanelWithFilter() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Помещения на этаже", SPACE_PANEL_COLOR));

        // ===== Список помещений =====
        spaceList = new JList<>(spaceListModel);
        spaceList.setCellRenderer(new SpaceListRenderer());
        spaceList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceList.setDragEnabled(true);
        spaceList.setDropMode(DropMode.INSERT);
        spaceList.setTransferHandler(spaceReorderHandler);
        registerDeleteKeyAction(spaceList, () -> removeSpace(null));

        // ===== Фильтры =====
        JPanel filters = buildSpaceFilterPanel();

        panel.add(filters, BorderLayout.NORTH);
        panel.add(new JScrollPane(spaceList), BorderLayout.CENTER);
        panel.add(createSpaceButtons(), BorderLayout.SOUTH);
        return panel;
    }
    //** Верхняя панель с 3 переключателями-фильтрами. Любая комбинация допустима. */
    private JPanel buildSpaceFilterPanel() {
        JPanel wrap = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 6));

        filterApartmentBtn = new JToggleButton("Жилые");
        filterOfficeBtn    = new JToggleButton("Офисные");
        filterPublicBtn    = new JToggleButton("Общественные");
        filterStreetBtn    = new JToggleButton("Улица"); // НОВОЕ

        // По умолчанию показываем всё
        filterApartmentBtn.setSelected(true);
        filterOfficeBtn.setSelected(true);
        filterPublicBtn.setSelected(true);
        filterStreetBtn.setSelected(true); // НОВОЕ

        // Стиль
        java.util.List<JToggleButton> all = java.util.List.of(
                filterApartmentBtn, filterOfficeBtn, filterPublicBtn, filterStreetBtn
        );
        for (JToggleButton b : all) {
            b.setFocusPainted(false);
            b.setFont(new Font("Segoe UI", Font.PLAIN, 13));
            b.setEnabled(true); // никаких авто-отключений
            b.setToolTipText(null);
        }

        // Любое изменение тумблеров → пересобрать список этажей и помещений
        java.awt.event.ItemListener l = e -> {
            refreshFloorListForSelectedSection(); // прячем/показываем этажи по типам
            updateSpaceList();                    // обновим список помещений текущего этажа (без автологики)
        };
        for (JToggleButton b : all) b.addItemListener(l);

        wrap.add(new JLabel("Фильтр этажей:"));
        wrap.add(filterApartmentBtn);
        wrap.add(filterOfficeBtn);
        wrap.add(filterPublicBtn);
        wrap.add(filterStreetBtn); // НОВОЕ
        return wrap;
    }


    private <T> JPanel createListPanel(String title, Color color, ListModel<T> model,
                                       ListCellRenderer<? super T> renderer,
                                       Supplier<JPanel> buttonPanelSupplier) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder(title, color));

        JList<T> list = new JList<>(model);
        list.setCellRenderer(renderer);
        list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        // Сохраняем ссылки на списки + включаем DnD там, где нужно
        if ("Этажи здания".equals(title)) {
            floorList = (JList<Floor>) list;
            registerDeleteKeyAction(floorList, () -> removeFloor(null));
        } else if ("Помещения на этаже".equals(title)) {
            spaceList = (JList<Space>) list;
            spaceList.setDragEnabled(true);
            spaceList.setDropMode(DropMode.INSERT);
            spaceList.setTransferHandler(spaceReorderHandler);
            registerDeleteKeyAction(spaceList, () -> removeSpace(null));
        } else if ("Комнаты в помещении".equals(title)) {
            roomList = (JList<Room>) list;
            roomList.setDragEnabled(true);
            roomList.setDropMode(DropMode.INSERT);
            roomList.setTransferHandler(roomReorderHandler);
            registerDeleteKeyAction(roomList, () -> removeRoom(null));
        }

        panel.add(new JScrollPane(list), BorderLayout.CENTER);
        panel.add(buttonPanelSupplier.get(), BorderLayout.SOUTH);

        return panel;
    }


    private TitledBorder createTitledBorder(String title, Color color) {
        return BorderFactory.createTitledBorder(
                null, title, TitledBorder.LEFT, TitledBorder.TOP, HEADER_FONT, color);
    }

    private JPanel createFloorButtons() {
        return createButtonPanel(
                createStyledButton("", FontAwesomeSolid.LAYER_GROUP, new Color(63,81,181), this::manageSections),
                createStyledButton("", FontAwesomeSolid.COPY,        new Color(0,150,136), this::copySection), // ← НОВОЕ
                createStyledButton("", FontAwesomeSolid.PLUS,        new Color(46,125,50),  this::addFloor),
                createStyledButton("", FontAwesomeSolid.CLONE,       new Color(100,181,246), this::copyFloor),
                createStyledButton("", FontAwesomeSolid.EDIT,        new Color(255,152,0),  this::editFloor),
                createStyledButton("", FontAwesomeSolid.TRASH,       new Color(198,40,40),  this::removeFloor)
        );

    }

    private JPanel createSpaceButtons() {
        return createButtonPanel(
                createStyledButton("", FontAwesomeSolid.PLUS, new Color(46, 125, 50), this::addSpace),
                createStyledButton("", FontAwesomeSolid.CLONE, new Color(100, 181, 246), this::copySpace),
                createStyledButton("", FontAwesomeSolid.EDIT, new Color(255, 152, 0), this::editSpace),
                createStyledButton("", FontAwesomeSolid.TRASH, new Color(198, 40, 40), this::removeSpace)
        );
    }

    private JPanel createRoomButtons() {
        return createButtonPanel(
                createStyledButton("", FontAwesomeSolid.PLUS, new Color(46, 125, 50), this::addRoom),
                createStyledButton("", FontAwesomeSolid.EDIT, new Color(255, 152, 0), this::editRoom),
                createStyledButton("", FontAwesomeSolid.TRASH, new Color(198, 40, 40), this::removeRoom)
        );
    }

    private JPanel createButtonPanel(JButton... buttons) {
        JPanel panel = new JPanel(new GridLayout(1, buttons.length, BUTTON_PANEL_SIZE.width, BUTTON_PANEL_SIZE.height));
        for (JButton button : buttons) {
            panel.add(button);
        }
        return panel;
    }

    private void registerDeleteKeyAction(JList<?> list, Runnable deleteAction) {
        if (list == null || deleteAction == null) return;

        final Object actionKey = "deleteSelectedItem_" + System.identityHashCode(list);
        InputMap inputMap = list.getInputMap(JComponent.WHEN_FOCUSED);
        ActionMap actionMap = list.getActionMap();
        inputMap.put(KeyStroke.getKeyStroke(KeyEvent.VK_DELETE, 0), actionKey);
        actionMap.put(actionKey, new AbstractAction() {
            @Override
            public void actionPerformed(ActionEvent e) {
                deleteAction.run();
            }
        });
    }

    private JPanel createActionButtons() {
        JPanel wrap = new JPanel(new BorderLayout());
        final JFXPanel fx = new JFXPanel();
        wrap.add(fx, BorderLayout.CENTER);

        Platform.runLater(() -> {
            // Кнопки
            javafx.scene.control.Button btnLoad    = new javafx.scene.control.Button("Загрузить проект");
            javafx.scene.control.Button btnSave    = new javafx.scene.control.Button("Сохранить проект");
            javafx.scene.control.Button btnExport  = new javafx.scene.control.Button("Экспорт в файл");
            javafx.scene.control.Button btnImport  = new javafx.scene.control.Button("Импорт из файла");
            javafx.scene.control.Button btnCalc    = new javafx.scene.control.Button("Рассчитать показатели");
            javafx.scene.control.Button btnSummary = new javafx.scene.control.Button("Сводка квартир");

            // Иконки (Ikonli JavaFX)
            org.kordamp.ikonli.javafx.FontIcon icLoad    = new org.kordamp.ikonli.javafx.FontIcon("fas-folder-open");
            org.kordamp.ikonli.javafx.FontIcon icSave    = new org.kordamp.ikonli.javafx.FontIcon("fas-save");
            org.kordamp.ikonli.javafx.FontIcon icExport  = new org.kordamp.ikonli.javafx.FontIcon("fas-file-export");
            org.kordamp.ikonli.javafx.FontIcon icImport  = new org.kordamp.ikonli.javafx.FontIcon("fas-file-import");
            org.kordamp.ikonli.javafx.FontIcon icCalc    = new org.kordamp.ikonli.javafx.FontIcon("fas-calculator");
            org.kordamp.ikonli.javafx.FontIcon icSummary = new org.kordamp.ikonli.javafx.FontIcon("fas-table");
            icLoad.setIconSize(16);
            icSave.setIconSize(16);
            icExport.setIconSize(16);
            icImport.setIconSize(16);
            icCalc.setIconSize(16);
            icSummary.setIconSize(16);
            btnLoad.setGraphic(icLoad);
            btnSave.setGraphic(icSave);
            btnExport.setGraphic(icExport);
            btnImport.setGraphic(icImport);
            btnCalc.setGraphic(icCalc);
            btnSummary.setGraphic(icSummary);

            // CSS-классы для цветов/hover
            btnLoad.getStyleClass().addAll("button", "btn-load");
            btnSave.getStyleClass().addAll("button", "btn-save");
            btnExport.getStyleClass().addAll("button", "btn-save");
            btnImport.getStyleClass().addAll("button", "btn-load");
            btnCalc.getStyleClass().addAll("button", "btn-calc");
            btnSummary.getStyleClass().addAll("button", "btn-summary");

            // Растягиваем равномерно
            javafx.scene.layout.HBox box = new javafx.scene.layout.HBox(10, btnLoad, btnSave, btnExport, btnImport, btnCalc, btnSummary);
            box.getStyleClass().addAll("controls-bar", "theme-light");
            for (javafx.scene.control.Button b : java.util.List.of(btnLoad, btnSave, btnExport, btnImport, btnCalc, btnSummary)) {
                b.setMaxWidth(Double.MAX_VALUE);
                javafx.scene.layout.HBox.setHgrow(b, javafx.scene.layout.Priority.ALWAYS);
            }

            // Сцена + CSS
            javafx.scene.Scene scene = new javafx.scene.Scene(box);
            scene.getStylesheets().add(
                    java.util.Objects.requireNonNull(getClass().getResource("/ui/protokol.css")).toExternalForm()
            );
            fx.setScene(scene);

            // Горячая смена темы (Ctrl+T)
            scene.setOnKeyPressed(ke -> {
                if (ke.isControlDown() && ke.getCode() == javafx.scene.input.KeyCode.T) {
                    java.util.List<String> cls = box.getStyleClass();
                    if (cls.contains("theme-light")) {
                        cls.remove("theme-light");
                        cls.add("theme-dark");
                    } else {
                        cls.remove("theme-dark");
                        cls.add("theme-light");
                    }
                }
            });

            // Цвета и "активность"
            final String COL_LOAD    = "#3949ab";
            final String COL_SAVE    = "#43a047";
            final String COL_EXPORT  = "#00897b";
            final String COL_IMPORT  = "#6d4c41";
            final String COL_CALC    = "#1e88e5";
            final String COL_SUMMARY = "#00897b";

            final java.util.List<javafx.scene.control.Button> all =
                    java.util.List.of(btnLoad, btnSave, btnExport, btnImport, btnCalc, btnSummary);

            final java.util.function.BiConsumer<javafx.scene.control.Button, String> markActive =
                    (btn, hex) -> {
                        all.forEach(ru.citlab24.protokol.ui.fx.ActiveGlow::clear);
                        // неон на ~50%
                        ru.citlab24.protokol.ui.fx.ActiveGlow.setActive(btn, hex, 0.5);
                    };

            // Обработчики: сначала отметить активной, потом вызвать Swing-логику
            btnLoad.setOnAction(ev -> {
                markActive.accept(btnLoad, COL_LOAD);
                javax.swing.SwingUtilities.invokeLater(() -> loadProject(null));
            });
            btnSave.setOnAction(ev -> {
                markActive.accept(btnSave, COL_SAVE);
                javax.swing.SwingUtilities.invokeLater(() -> saveProject(null));
            });
            btnExport.setOnAction(ev -> {
                markActive.accept(btnExport, COL_EXPORT);
                javax.swing.SwingUtilities.invokeLater(() -> exportProject(null));
            });
            btnImport.setOnAction(ev -> {
                markActive.accept(btnImport, COL_IMPORT);
                javax.swing.SwingUtilities.invokeLater(() -> importProject(null));
            });
            btnCalc.setOnAction(ev -> {
                markActive.accept(btnCalc, COL_CALC);
                javax.swing.SwingUtilities.invokeLater(() -> calculateMetrics(null));
            });
            btnSummary.setOnAction(ev -> {
                markActive.accept(btnSummary, COL_SUMMARY);
                javax.swing.SwingUtilities.invokeLater(() -> showApartmentSummary(null));
            });
        });

        return wrap;
    }

    private void copySection(ActionEvent e) {
        if (sectionList == null || sectionList.isSelectionEmpty()) {
            showMessage("Выберите секцию для копирования", "Информация", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        // 0) Сначала зафиксируем состояния из вкладок в модель
        MicroclimateTab microTab = getMicroclimateTab();
        if (microTab != null) microTab.updateRoomSelectionStates(); // ← МИКРОКЛИМАТ в Room

        // 1) Сохраняем «снимки» галочек
        Map<String, Boolean> savedMicroSelections = saveMicroclimateSelections(); // МИКРОКЛИМАТ
        Map<String, Boolean> savedRadSelections   = saveRadiationSelections();    // РАДИАЦИЯ (как было)

        int srcIdx = sectionList.getSelectedIndex();
        Section src = sectionListModel.get(srcIdx);

        String suggested = generateUniqueSectionName(src.getName());
        String newName = showInputDialog("Копировать секцию", "Название новой секции:", suggested);
        if (newName == null || newName.trim().isEmpty()) return;

        // создаём новую секцию в конце списка
        int newPosition = building.getSections().size();
        Section dst = new Section();
        dst.setName(newName.trim());
        dst.setPosition(newPosition);
        building.addSection(dst);

        int dstIdx = newPosition;

        // 2) Глубоко копируем этажи исходной секции (rooms с НОВЫМИ id)
        java.util.List<Floor> toAdd = new java.util.ArrayList<>();
        for (Floor f : building.getFloors()) {
            if (f.getSectionIndex() == srcIdx) {
                Floor fCopy = createFloorCopy(f);
                fCopy.setSectionIndex(dstIdx);
                toAdd.add(fCopy);
            }
        }
        for (Floor fCopy : toAdd) building.addFloor(fCopy);

        // 3) ВОССТАНОВЛЕНИЕ МИКРОКЛИМАТА: применяем снимок к ВСЕМ комнатам здания
        restoreMicroclimateSelections(savedMicroSelections);

        // 4) UI
        refreshSectionListModel();
        sectionList.setSelectedIndex(dstIdx);
        refreshFloorListForSelectedSection();
        updateSpaceList();
        updateRoomList();

        // 5) Обновляем вкладки (порядок важен)
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        restoreRadiationSelections(savedRadSelections);

        updateLightingTab(building, /*autoApplyDefaults=*/true);

        if (microTab != null) {
            try { microTab.selectSectionByIndex(dstIdx); } catch (Throwable ignore) {}
        }

        showMessage("Секция «" + src.getName() + "» успешно скопирована в «" + newName + "».",
                "Готово", JOptionPane.INFORMATION_MESSAGE);
    }

    private String generateUniqueSectionName(String base) {
        if (base == null || base.isEmpty()) base = "Секция";
        String plain = base + " (копия)";
        java.util.Set<String> names = new java.util.HashSet<>();
        for (Section s : building.getSections()) {
            if (s.getName() != null) names.add(s.getName());
        }
        if (!names.contains(plain)) return plain;

        // если имя занято — добавляем счётчик
        int n = 2;
        while (names.contains(plain + " " + n)) n++;
        return plain + " " + n;
    }

    private JPanel createSectionsAndFloorsPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Секции и этажи", FLOOR_PANEL_COLOR));

        // СЛЕВА — список секций
        sectionList = new JList<>(sectionListModel);
        sectionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        sectionList.setFixedCellHeight(28);
        refreshSectionListModel();

// СПРАВА — список этажей текущей секции
        floorList = new JList<>(floorListModel);
        floorList.setCellRenderer(new FloorListRenderer());
        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        floorList.setFixedCellHeight(28);
        registerDeleteKeyAction(floorList, () -> removeFloor(null));

        if (!sectionListModel.isEmpty()) sectionList.setSelectedIndex(0);
        refreshFloorListForSelectedSection();

        sectionList.setDragEnabled(true);
        sectionList.setDropMode(DropMode.INSERT);
        sectionList.setTransferHandler(sectionReorderHandler);

        floorList.setDragEnabled(true);
        floorList.setDropMode(DropMode.INSERT);
        floorList.setTransferHandler(floorReorderHandler);
        floorList.setDragEnabled(true);
        floorList.setDropMode(DropMode.INSERT);
        floorList.setTransferHandler(floorReorderHandler);

        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) refreshFloorListForSelectedSection();
        });

        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
                new JScrollPane(sectionList), new JScrollPane(floorList));
        split.setResizeWeight(0.35);
        panel.add(split, BorderLayout.CENTER);
        panel.add(createFloorButtons(), BorderLayout.SOUTH);
        return panel;
    }
    // Флажки, чтобы понимать откуда тянем (этажи → на секцию, или секции ↔ секции)
    private boolean draggingFromFloor = false;

    // Перестановка этажей ВНУТРИ текущей секции
    private final TransferHandler floorReorderHandler = new ReorderHandler<Floor>() {
        @Override protected Transferable createTransferable(JComponent c) {
            draggingFromFloor = true;                    // помечаем источник DnD
            return super.createTransferable(c);
        }
        @Override protected void afterReorder(DefaultListModel<Floor> model) {
            // Переименовываем позиции по текущему порядку в списке
            for (int i = 0; i < model.size(); i++) {
                model.get(i).setPosition(i);
            }
            draggingFromFloor = false;
        }
    };

    // Перестановка секций МЕЖДУ собой (ПЕРЕНОС ЭТАЖЕЙ НА СЕКЦИЮ ЗАПРЕЩЁН)
    private final TransferHandler sectionReorderHandler = new ReorderHandler<Section>() {

        @Override
        public boolean canImport(TransferSupport s) {
            // Запрещаем бросать ЭТАЖИ на список секций (перенос между секциями)
            if (draggingFromFloor) return false;
            return super.canImport(s);
        }

        @Override
        public boolean importData(TransferSupport s) {
            // Перенос этажа на секцию запрещён
            if (draggingFromFloor) return false;
            // Разрешаем только перестановку СЕКЦИЙ между собой
            return super.importData(s);
        }

        @Override
        protected void afterReorder(DefaultListModel<Section> model) {
            // 1) Текущий выбор секции (чтобы восстановить выделение)
            Section selected = (sectionList != null) ? sectionList.getSelectedValue() : null;

            // 2) Снимок старого списка секций (для ремапа индексов этажей)
            java.util.List<Section> oldSections = new java.util.ArrayList<>(building.getSections());

            // 3) Собираем новый порядок секций из модели JList и обновляем position
            java.util.List<Section> newSections = new java.util.ArrayList<>();
            for (int i = 0; i < model.size(); i++) {
                Section s = model.get(i);
                s.setPosition(i);
                newSections.add(s);
            }
            // 4) Карта: ссылка на секцию → новый индекс
            java.util.Map<Section, Integer> newIndexByRef = new java.util.HashMap<>();
            for (int i = 0; i < newSections.size(); i++) {
                newIndexByRef.put(newSections.get(i), i);
            }
            // 5) Ремапим sectionIndex у этажей так, чтобы они остались в СВОИХ секциях
            for (Floor f : building.getFloors()) {
                int oldIdx = f.getSectionIndex();
                if (oldIdx >= 0 && oldIdx < oldSections.size()) {
                    Section oldSec = oldSections.get(oldIdx);
                    Integer ni = newIndexByRef.get(oldSec);
                    f.setSectionIndex((ni != null) ? ni : 0);
                } else {
                    f.setSectionIndex(0);
                }
            }
            // 6) Фиксируем новые секции в модели здания и обновляем UI
            building.setSections(newSections);
            refreshSectionListModel();
            if (selected != null) {
                int selIdx = building.getSections().indexOf(selected);
                sectionList.setSelectedIndex(selIdx >= 0 ? selIdx : 0);
            } else if (!sectionListModel.isEmpty()) {
                sectionList.setSelectedIndex(0);
            }
            refreshFloorListForSelectedSection();
        }
    };

    private void refreshSectionListModel() {
        sectionListModel.clear();
        for (Section s : building.getSections()) sectionListModel.addElement(s);
    }

    private void refreshFloorListForSelectedSection() {
        floorListModel.clear();
        if (sectionList == null) return;

        int secIdx = Math.max(0, sectionList.getSelectedIndex());

        // Снимем текущий выбор, чтобы попытаться восстановить
        Floor previouslySelected = (floorList != null) ? floorList.getSelectedValue() : null;

        java.util.List<Floor> all = new java.util.ArrayList<>();
        for (Floor f : building.getFloors()) {
            if (f.getSectionIndex() == secIdx) all.add(f);
        }
        all.sort(java.util.Comparator.comparingInt(Floor::getPosition));

// Чёткая фильтрация по ТИПУ ЭТАЖА (Жилые/Офисные/Общественные/Улица)
        for (Floor f : all) {
            if (isFloorVisibleByFilter(f)) {
                floorListModel.addElement(f);
            }
        }

        if (!floorListModel.isEmpty()) {
            int idx = (previouslySelected != null) ? floorListModel.indexOf(previouslySelected) : -1;
            floorList.setSelectedIndex(idx >= 0 ? idx : 0);
        } else {
            if (spaceListModel != null) spaceListModel.clear();
            if (roomListModel  != null) roomListModel.clear();
        }

    }

    private JButton createStyledButton(String text, FontAwesomeSolid icon, java.awt.Color bgColor, ActionListener action) {
        JButton btn = new JButton(text, FontIcon.of(icon, 16, java.awt.Color.WHITE));
        btn.setBackground(bgColor);
        btn.setForeground(Color.WHITE);
        btn.setFocusPainted(false);
        btn.setFont(UIManager.getFont("Button.font"));
        btn.addActionListener(action);

        btn.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                btn.setBackground(bgColor.darker());
            }

            @Override
            public void mouseExited(MouseEvent e) {
                btn.setBackground(bgColor);
            }
        });
        return btn;
    }

    void manageSections(ActionEvent e) {
        JFrame frame = (JFrame) SwingUtilities.getWindowAncestor(this);
        ManageSectionsDialog dlg = new ManageSectionsDialog(frame, building.getSections());
        if (!dlg.showDialog()) return;

        // 1) Запоминаем старые секции (для ремапа этажей по имени)
        List<Section> oldSections = new ArrayList<>(building.getSections());
        List<Section> updated = dlg.getSections();

        // 2) Строим карту "имя секции -> новый индекс"
        Map<String, Integer> nameToNewIndex = new HashMap<>();
        for (int i = 0; i < updated.size(); i++) {
            String name = updated.get(i).getName();
            if (name != null) nameToNewIndex.put(name, i);
        }

        // 3) Ремапим sectionIndex у всех этажей по ИМЕНИ старой секции.
        // Если имя не найдено в новом списке — отправляем этаж в секцию 0.
        for (Floor f : building.getFloors()) {
            int oldIdx = f.getSectionIndex();
            String oldName = (oldIdx >= 0 && oldIdx < oldSections.size())
                    ? oldSections.get(oldIdx).getName()
                    : null;
            Integer newIdx = (oldName == null) ? null : nameToNewIndex.get(oldName);
            f.setSectionIndex(newIdx != null ? newIdx : 0);
        }

        // 4) Фиксируем новые секции и обновляем UI
        building.setSections(updated);
        refreshSectionListModel();
        if (!sectionListModel.isEmpty()) sectionList.setSelectedIndex(0);
        refreshFloorListForSelectedSection();
        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        updateNoiseTabNow();
    }

    /** Возвращает true, если этаж виден при текущих тумблерах. */
    private boolean isFloorVisibleByFilter(Floor f) {
        if (f == null) return true; // на всякий
        boolean ready = (filterApartmentBtn != null && filterOfficeBtn != null
                && filterPublicBtn != null && filterStreetBtn != null);

        boolean showA = !ready || filterApartmentBtn.isSelected();
        boolean showO = !ready || filterOfficeBtn.isSelected();
        boolean showP = !ready || filterPublicBtn.isSelected();
        boolean showS = !ready || filterStreetBtn.isSelected();

        Floor.FloorType t = f.getType();
        if (t == null) return (showA || showO || showP || showS);

        switch (t) {
            case RESIDENTIAL: return showA;
            case OFFICE:      return showO;
            case PUBLIC:      return showP;
            case STREET:      return showS;
            case MIXED:       return (showA || showO || showP); // смешанный видим, если включено хоть что-то из внутренних типов
            default:          return (showA || showO || showP || showS);
        }
    }

    // Основные операции с проектом
    private void loadProject(ActionEvent e) {
        try {
            List<Building> projects = DatabaseManager.getAllBuildings();
            if (projects.isEmpty()) {
                showMessage("Нет сохраненных проектов", "Информация", JOptionPane.INFORMATION_MESSAGE);
                return;
            }

            LoadProjectDialog dialog = new LoadProjectDialog(
                    (JFrame) SwingUtilities.getWindowAncestor(this), projects);
            dialog.setVisible(true);

            Building selectedProject = dialog.getSelectedProject();
            if (selectedProject != null) {
                loadSelectedProject(selectedProject);
            }
        } catch (SQLException ex) {
            handleError("Ошибка загрузки проектов: " + ex.getMessage(), "Ошибка");
        }
    }

    private void exportProject(ActionEvent e) {
        try {
            int currentId = (building != null) ? building.getId() : 0;
            String currentName = (projectNameField != null) ? projectNameField.getText() : null;
            ProjectFileService.exportProject(this, currentId, currentName);
        } catch (Exception ex) {
            handleError("Не удалось экспортировать проект: " + ex.getMessage(), "Ошибка");
        }
    }

    private void importProject(ActionEvent e) {
        try {
            Building imported = ProjectFileService.importProject(this);
            if (imported != null) {
                loadSelectedProject(imported);
            }
        } catch (Exception ex) {
            handleError("Не удалось импортировать проект: " + ex.getMessage(), "Ошибка");
        }
    }

    private void loadSelectedProject(Building selectedProject) throws SQLException {
        Building loadedBuilding = DatabaseManager.loadBuilding(selectedProject.getId());
        this.building = loadedBuilding;
        this.ops.setBuilding(this.building);
        projectNameField.setText(loadedBuilding.getName());

        // Обновляем списки/вкладки
        refreshAllLists();
        updateVentilationTab(loadedBuilding);
        updateRadiationTab(loadedBuilding, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(loadedBuilding, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(loadedBuilding, /*autoApplyDefaults=*/false);

        // Искусственное освещение — галочки из БД
        ArtificialLightingTab alt = getArtificialLightingTab();
        if (alt != null) {
            try {
                Map<String, Boolean> fromDb =
                        DatabaseManager.loadArtificialSelectionsByKey(loadedBuilding.getId());
                alt.applySelectionsByKey(loadedBuilding, fromDb);
                alt.refreshData();
            } catch (SQLException ex) {
                handleError("Не удалось загрузить галочки искусственного освещения: " + ex.getMessage(), "Ошибка");
            }
        }

        // Осв улица — значения из БД
        try {
            java.util.Map<String, Double[]> streetVals =
                    DatabaseManager.loadStreetLightingValuesByKey(loadedBuilding.getId());
            StreetLightingTab street = getStreetLightingTab();
            if (street != null) {
                street.setBuilding(loadedBuilding);
                street.refreshData();
                street.applyValuesByKey(streetVals);
            }
        } catch (SQLException ex) {
            handleError("Не удалось загрузить значения 'Осв улица': " + ex.getMessage(), "Ошибка");
        }

        // НОВОЕ: «Шумы» — применить сохранённые настройки
        try {
            ru.citlab24.protokol.tabs.modules.noise.NoiseTab noise = getNoiseTab();
            if (noise != null) {
                java.util.Map<String, DatabaseManager.NoiseValue> nv =
                        DatabaseManager.loadNoiseSelectionsByKey(loadedBuilding.getId());
                java.util.Map<String, double[]> th =
                        DatabaseManager.loadNoiseThresholds(loadedBuilding.getId());
                noise.setBuilding(loadedBuilding);
                noise.applySelectionsByKey(nv);
                noise.applyThresholds(th);
                noise.refreshData();
            }
        } catch (SQLException ex) {
            handleError("Не удалось загрузить настройки 'Шумы': " + ex.getMessage(), "Ошибка");
        }

        showMessage("Проект '" + loadedBuilding.getName() + "' успешно загружен",
                "Загрузка", JOptionPane.INFORMATION_MESSAGE);
    }

    private void saveProject(ActionEvent e) {
        logger.info("BuildingTab.saveProject() - Начало сохранения проекта");

        // 0) Финализируем активное редактирование таблиц
        try {
            KeyboardFocusManager kfm = KeyboardFocusManager.getCurrentKeyboardFocusManager();
            Component fo = (kfm != null) ? kfm.getFocusOwner() : null;
            JTable editingTable = (fo == null) ? null
                    : (JTable) SwingUtilities.getAncestorOfClass(JTable.class, fo);
            if (editingTable != null && editingTable.isEditing()) {
                try { editingTable.getCellEditor().stopCellEditing(); } catch (Exception ignore) {}
            }
        } catch (Exception ignore) {}

        // 1) Синхронизируем вкладки → модель
        RadiationTab radiationTab = getRadiationTab();
        if (radiationTab != null) {
            radiationTab.updateRoomSelectionStates();
        }

        LightingTab lightingTab = getLightingTab();
        if (lightingTab != null) {
            lightingTab.updateRoomSelectionStates();
        }

        // КЕО: снимок
        Map<String, Boolean> snapKeo = saveKeoSelections();

        // Искусственное: снимок
        ArtificialLightingTab artificialTab = getArtificialLightingTab();
        Map<String, Boolean> snapArtificial = java.util.Collections.emptyMap();
        if (artificialTab != null) {
            artificialTab.updateRoomSelectionStates();
            snapArtificial = artificialTab.saveSelectionsByKey();
        }

        // Осв улица — снимок
        StreetLightingTab street = getStreetLightingTab();
        java.util.Map<String, Double[]> snapStreet = java.util.Collections.emptyMap();
        if (street != null) {
            snapStreet = street.snapshotValuesByKey();
        }

        // НОВОЕ: Шумы — снимок по ключу
        ru.citlab24.protokol.tabs.modules.noise.NoiseTab noise = getNoiseTab();
        java.util.Map<String, ru.citlab24.protokol.db.DatabaseManager.NoiseValue> snapNoise = java.util.Collections.emptyMap();
        java.util.Map<String, double[]> snapNoiseThresholds = java.util.Collections.emptyMap();
        if (noise != null) {
            noise.updateRoomSelectionStates();
            snapNoise = noise.saveSelectionsByKey();
            snapNoiseThresholds = noise.saveThresholdsByKey();
        }

        // Микроклимат
        MicroclimateTab microTab = getMicroclimateTab();
        if (microTab != null) microTab.updateRoomSelectionStates();

        // 2) Имя проекта
        String baseName = projectNameField.getText().trim();
        if (baseName.isEmpty()) {
            showMessage("Введите название проекта!", "Ошибка", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 3) Копия проекта
        Building newProject = createBuildingCopy();
        newProject.setName(generateProjectVersionName(baseName));

        // 3.1) Применяем снимок КЕО к копии до сохранения
        restoreKeoSelections(newProject, snapKeo);

        // 4) Сохранение в БД
        try {
            DatabaseManager.saveBuilding(newProject);
        } catch (SQLException ex) {
            handleError("Ошибка сохранения: " + ex.getMessage(), "Ошибка");
            return;
        }

        // 4.1) Искусственное — записать галочки в новые room.id
        try {
            DatabaseManager.updateArtificialSelections(newProject, snapArtificial);
        } catch (SQLException ex) {
            handleError("Не удалось сохранить галочки искусственного освещения: " + ex.getMessage(), "Ошибка");
        }

        // 4.2) Осв улица — записать 4 значения
        try {
            DatabaseManager.updateStreetLightingValues(newProject, snapStreet);
        } catch (SQLException ex) {
            handleError("Не удалось сохранить значения 'Осв улица': " + ex.getMessage(), "Ошибка");
        }

        // 4.3) НОВОЕ: Шумы — записать состояния/источники
        try {
            DatabaseManager.updateNoiseSelections(newProject, snapNoise);
        } catch (SQLException ex) {
            handleError("Не удалось сохранить настройки 'Шумы': " + ex.getMessage(), "Ошибка");
        }

        try {
            DatabaseManager.updateNoiseThresholds(newProject, snapNoiseThresholds);
        } catch (SQLException ex) {
            handleError("Не удалось сохранить пороги 'Шумы': " + ex.getMessage(), "Ошибка");
        }

        // 4.4) Синхронизация состояния UI
        this.building = newProject;
        this.ops.setBuilding(this.building);
        projectNameField.setText(extractBaseName(newProject.getName()));

        // 5) Обновляем вкладки (без автодефолтов)
        refreshAllLists();
        updateRadiationTab(newProject, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(newProject, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(newProject, /*autoApplyDefaults=*/false);

        // 5.1) Вернуть галочки искусственного во вкладку
        ArtificialLightingTab alt = getArtificialLightingTab();
        if (alt != null) {
            alt.applySelectionsByKey(newProject, snapArtificial);
            alt.refreshData();
        }

        // 5.2) Осв улица — вернуть значения
        StreetLightingTab st = getStreetLightingTab();
        if (st != null) {
            st.setBuilding(newProject);
            st.refreshData();
            st.applyValuesByKey(snapStreet);
        }

        // 5.3) НОВОЕ: Шумы — вернуть состояния
        if (noise != null) {
            noise.setBuilding(newProject);
            noise.applySelectionsByKey(snapNoise);
            noise.applyThresholds(snapNoiseThresholds);
            noise.refreshData();
        }

        logger.info("Проект успешно сохранен");
    }

    private String generateProjectVersionName(String baseName) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
        String currentDate = dateFormat.format(new Date());
        String cleanBaseName = extractBaseName(baseName);

        int nextVersion = calculateNextVersion(cleanBaseName);
        return nextVersion == 1 ?
                cleanBaseName :
                cleanBaseName + " ред." + nextVersion + " " + currentDate;
    }

    private int calculateNextVersion(String cleanBaseName) {
        try {
            List<Building> existingProjects = DatabaseManager.getAllBuildings();
            Pattern versionPattern = Pattern.compile("^" + Pattern.quote(cleanBaseName) + "(?: ред\\.(\\d+) (\\d{2}\\.\\d{2}\\.\\d{4}))?$");

            int maxVersion = existingProjects.stream()
                    .map(project -> versionPattern.matcher(project.getName()))
                    .filter(Matcher::find)
                    .mapToInt(matcher -> matcher.group(1) != null ? Integer.parseInt(matcher.group(1)) : 1)
                    .max()
                    .orElse(0);

            return maxVersion + 1;
        } catch (SQLException e) {
            handleError("Ошибка получения списка проектов: " + e.getMessage(), "Ошибка");
            return 1;
        }
    }

    private String extractBaseName(String name) {
        Pattern pattern = Pattern.compile("(.+?) (?:ред\\.\\d+ \\d{2}\\.\\d{2}\\.\\d{4})$");
        Matcher matcher = pattern.matcher(name);
        return matcher.find() ? matcher.group(1).trim() : name;
    }

    private Building createBuildingCopy() {
        Building copy = new Building();
        copy.setName(building.getName());

        // копируем секции
        List<Section> copiedSections = new ArrayList<>();
        for (Section s : building.getSections()) {
            Section cs = new Section();
            cs.setName(s.getName());
            cs.setPosition(s.getPosition());
            copiedSections.add(cs);
        }
        copy.setSections(copiedSections);

        // этажи — с сохранением selected и id
        for (Floor originalFloor : building.getFloors()) {
            Floor floorCopy = createFloorCopyPreserve(originalFloor);
            copy.addFloor(floorCopy);
        }
        return copy;
    }

    @SuppressWarnings("serial")
    private abstract class ReorderHandler<T> extends TransferHandler {
        private final DataFlavor flavor = DataFlavor.stringFlavor;
        @Override protected Transferable createTransferable(JComponent c) {
            @SuppressWarnings("unchecked") JList<T> list = (JList<T>) c;
            return new StringSelection(Integer.toString(list.getSelectedIndex()));
        }
        @Override public int getSourceActions(JComponent c) { return MOVE; }
        @Override public boolean canImport(TransferSupport s) { return s.isDrop() && s.isDataFlavorSupported(flavor); }
        @Override public boolean importData(TransferSupport s) {
            try {
                @SuppressWarnings("unchecked") JList<T> list = (JList<T>) s.getComponent();
                DefaultListModel<T> model = (DefaultListModel<T>) list.getModel();
                int to = ((JList.DropLocation) s.getDropLocation()).getIndex();
                int from = Integer.parseInt((String) s.getTransferable().getTransferData(flavor));
                if (from == -1 || from == to) return false;
                T elem = model.get(from);
                model.remove(from);
                if (to > from) to--;
                model.add(to, elem);
                list.setSelectedIndex(to);
                afterReorder(model);
                return true;
            } catch (Exception ex) { ex.printStackTrace(); return false; }
        }
        protected abstract void afterReorder(DefaultListModel<T> model);
    }

    private final TransferHandler spaceReorderHandler = new ReorderHandler<Space>() {
        @Override protected void afterReorder(DefaultListModel<Space> model) {
            // Переписываем position по новому порядку
            for (int i = 0; i < model.size(); i++) model.get(i).setPosition(i);
            updateLightingTab(building, /*autoApplyDefaults=*/true);
            updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        }
    };

    private final TransferHandler roomReorderHandler = new ReorderHandler<Room>() {
        @Override protected void afterReorder(DefaultListModel<Room> model) {
            for (int i = 0; i < model.size(); i++) model.get(i).setPosition(i);
            updateLightingTab(building, /*autoApplyDefaults=*/true);
            updateMicroclimateTab(building, /*autoApplyDefaults=*/false);


        }
    };

    private Floor createFloorCopy(Floor originalFloor) { return ops.createFloorCopy(originalFloor); }

    private void calculateMetrics(ActionEvent e) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            // 1) Пересчёт вентиляции
            updateVentilationTab(building);

            // 2) Радиация: НЕ трогаем старые галочки
            RadiationTab rt = getRadiationTab();
            Map<String, Boolean> snap = (rt != null) ? rt.saveSelections()
                    : java.util.Collections.emptyMap();

            // Перерисовали без forceOfficeSelection и без авто-правил
            updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);

            // Вернули снимок
            if (rt != null) rt.restoreSelections(snap);

            // Переключаемся на вкладку вент.
            ((MainFrame) mainFrame).selectVentilationTab();
        }
        showMessage("Данные для вентиляции обновлены!", "Расчет завершен", JOptionPane.INFORMATION_MESSAGE);
    }

    private void showApartmentSummary(ActionEvent e) {
        // Безопасность: если вдруг building == null, создаём пустой
        if (this.building == null) this.building = new Building();
        ApartmentSummaryDialog.show(this, this.building);
    }

    // Операции с помещениями
    private void copySpace(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        Space selectedSpace = spaceList.getSelectedValue();

        if (selectedFloor == null || selectedSpace == null) {
            showMessage("Выберите помещение для копирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Сохраняем состояния чекбоксов ДО копирования (радиация)
        Map<String, Boolean> savedSelections = saveRadiationSelections();

        // Создаём копию помещения с НОВЫМИ id комнат
        Space copiedSpace = createSpaceCopyWithNewIds(selectedSpace);
        copiedSpace.setIdentifier(generateNextSpaceId(selectedFloor, selectedSpace.getIdentifier()));

        // позиция в КОНЕЦ
        int maxPos = selectedFloor.getSpaces().stream()
                .mapToInt(Space::getPosition)
                .max().orElse(-1);
        copiedSpace.setPosition(maxPos + 1);


        // Добавляем
        selectedFloor.addSpace(copiedSpace);

        // Правильно пересобираем UI-список
        updateSpaceList();
        spaceList.setSelectedValue(copiedSpace, true);

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        restoreRadiationSelections(savedSelections);  // вернули старые состояния

        RadiationTab rt = getRadiationTab();
        if (rt != null) {
            // Если на ЭТОМ этаже уже есть помещения с галочками — у КОПИИ все галочки СНИМАЕМ
            if (rt.hasAnySelectedOnFloor(selectedFloor)) {
                rt.clearSelectionsForSpace(copiedSpace);
            } else {
                // Этаж «чистый» — можно применить дефолты к новой копии
                rt.applyDefaultsForSpace(copiedSpace);
            }

            int newIndex = selectedFloor.getSpaces().indexOf(copiedSpace);
            if (newIndex >= 0) rt.selectSpaceByIndex(newIndex);
        }
    }

    private Space createSpaceCopyWithNewIds(Space originalSpace) { return ops.createSpaceCopyWithNewIds(originalSpace); }

    private String generateNextSpaceId(Floor floor, String currentId) {
        String prefix = "";
        String baseNumber = "";
        int lastDashIndex = currentId.lastIndexOf('-');

        if (lastDashIndex != -1) {
            prefix = currentId.substring(0, lastDashIndex).trim();
            baseNumber = currentId.substring(lastDashIndex + 1).trim();
        } else {
            prefix = currentId;
        }

        int currentNumber = parseNumber(baseNumber);
        int maxNumber = calculateMaxSpaceNumber(floor, prefix, currentNumber);
        int nextNumber = maxNumber + 1;

        return prefix.isEmpty() ?
                String.valueOf(nextNumber) :
                prefix + "-" + nextNumber;
    }

    private int parseNumber(String numberStr) {
        try {
            String numericPart = numberStr.replaceAll("\\D", "");
            return numericPart.isEmpty() ? 0 : Integer.parseInt(numericPart);
        } catch (NumberFormatException ignored) {
            return 0;
        }
    }

    private int calculateMaxSpaceNumber(Floor floor, String prefix, int currentNumber) {
        int maxNumber = currentNumber;
        for (Space space : floor.getSpaces()) {
            String id = space.getIdentifier();
            if (id.startsWith(prefix)) {
                String suffix = id.substring(prefix.length()).trim();
                if (suffix.startsWith("-")) suffix = suffix.substring(1).trim();

                int num = parseNumber(suffix);
                if (num > maxNumber) maxNumber = num;
            }
        }
        return maxNumber;
    }
    private void addFloor(ActionEvent e) {
        AddFloorDialog dialog = new AddFloorDialog((JFrame) SwingUtilities.getWindowAncestor(this));
        if (dialog.showDialog()) {
            Floor floor = new Floor();
            floor.setNumber(dialog.getFloorNumber());
            floor.setType(dialog.getFloorType());

            int secIdx = Math.max(0, sectionList.getSelectedIndex());
            floor.setSectionIndex(secIdx); // ← секция

            String floorName = floor.getType().title + " " + floor.getNumber();
            floor.setName(floorName);

            int maxPos = building.getFloors().stream()
                    .filter(f -> f.getSectionIndex() == secIdx)
                    .mapToInt(Floor::getPosition)
                    .max().orElse(-1);
            floor.setPosition(maxPos + 1);

            building.addFloor(floor);

            refreshFloorListForSelectedSection();

            RadiationTab rt = getRadiationTab();
            Map<String, Boolean> snap = (rt != null) ? rt.saveSelections() : java.util.Collections.emptyMap();

            updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);

            if (rt != null) {
                rt.restoreSelections(snap);
                rt.applyDefaultsForFloorFirstResidentialOnly(floor);
            }

            updateLightingTab(building, /*autoApplyDefaults=*/true);
            updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

            // НОВОЕ
            updateNoiseTabNow();
        }
    }

    private void updateRadiationTab(Building building, boolean forceOfficeSelection, boolean autoApplyRules) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            RadiationTab tab = ((MainFrame) mainFrame).getRadiationTab();
            if (tab != null) {
                tab.setBuilding(building, forceOfficeSelection, autoApplyRules);
            }
        }
    }

    private RadiationTab getRadiationTab() {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            return ((MainFrame) mainFrame).getRadiationTab();
        }
        return null;
    }

    private void copyFloor(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) return;

        RadiationTab radiationTab = getRadiationTab();
        Map<Integer, Boolean> allRoomStates = new HashMap<>();
        if (radiationTab != null) {
            allRoomStates.putAll(radiationTab.globalRoomSelectionMap);
        }
        LightingTab lightingTab = getLightingTab();
        if (lightingTab != null) {
            lightingTab.updateRoomSelectionStates();
        }

        Floor copiedFloor = createFloorCopy(selectedFloor);
        String newFloorNumber = generateNextFloorNumber(selectedFloor.getNumber(), selectedFloor.getSectionIndex());
        copiedFloor.setNumber(newFloorNumber);
        copiedFloor.setSectionIndex(selectedFloor.getSectionIndex());

        int secIdx = selectedFloor.getSectionIndex();
        int maxPos = building.getFloors().stream()
                .filter(f -> f.getSectionIndex() == secIdx)
                .mapToInt(Floor::getPosition)
                .max().orElse(-1);
        copiedFloor.setPosition(maxPos + 1);

        updateSpaceIdentifiers(copiedFloor, extractDigits(newFloorNumber));

        building.addFloor(copiedFloor);
        floorListModel.addElement(copiedFloor);

        floorList.setSelectedValue(copiedFloor, true);
        updateLightingTab(building, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections() : java.util.Collections.emptyMap();

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);

        if (rt != null) {
            rt.restoreSelections(snap);
            rt.applyDefaultsForFloorFirstResidentialOnly(copiedFloor);
            rt.refreshFloors();
        }
        updateNoiseTabNow();
    }

    private String extractDigits(String input) { return ops.extractDigits(input); }

    // Новый: считает следующий номер ТОЛЬКО в пределах указанной секции.
    private String generateNextFloorNumber(String currentNumber, int sectionIndex) {
        return ops.generateNextFloorNumber(currentNumber, sectionIndex);
    }

    private void updateSpaceList() {
        spaceListModel.clear();
        Floor selectedFloor = (floorList != null) ? floorList.getSelectedValue() : null;
        if (selectedFloor == null) return;

        java.util.List<Space> sorted = new java.util.ArrayList<>(selectedFloor.getSpaces());
        sorted.sort(java.util.Comparator.comparingInt(Space::getPosition));

        for (Space s : sorted) {
            if (isSpaceVisibleByFilterOnFloor(s, selectedFloor)) {
                spaceListModel.addElement(s); // ← на «улице» попадёт только OUTDOOR
            }
        }

        if (!spaceListModel.isEmpty()) {
            Space prev = spaceList.getSelectedValue();
            int idx = (prev != null) ? spaceListModel.indexOf(prev) : -1;
            spaceList.setSelectedIndex(idx >= 0 ? idx : 0);
        } else {
            roomListModel.clear();
        }
    }

    private void updateSpaceIdentifiers(Floor floor, String newFloorDigits) {
        ops.updateSpaceIdentifiers(floor, newFloorDigits);
    }

    private void editFloor(ActionEvent e) {
        int index = floorList.getSelectedIndex();
        if (index < 0) {
            showMessage("Выберите этаж для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Floor floor = floorListModel.get(index);
        AddFloorDialog dialog = new AddFloorDialog((JFrame) SwingUtilities.getWindowAncestor(this));
        dialog.setFloorNumber(floor.getNumber());
        dialog.setFloorType(floor.getType());

        if (dialog.showDialog()) {
            floor.setNumber(dialog.getFloorNumber());
            floor.setType(dialog.getFloorType());
            floorListModel.set(index, floor);
            updateSpaceList();
        }

        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections()
                : java.util.Collections.emptyMap();

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        if (rt != null) rt.restoreSelections(snap);

        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        updateNoiseTabNow();
    }

    private Map<String, Boolean> saveRadiationSelections() {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            RadiationTab tab = ((MainFrame) mainFrame).getRadiationTab();
            if (tab != null) {
                return tab.saveSelections();
            }
        }
        return new HashMap<>();
    }

    private void restoreRadiationSelections(Map<String, Boolean> savedSelections) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            RadiationTab tab = ((MainFrame) mainFrame).getRadiationTab();
            if (tab != null) {
                tab.restoreSelections(savedSelections);
            }
        }
    }

    private void removeFloor(ActionEvent e) {
        int index = floorList.getSelectedIndex();
        // Сначала сохраняем радиационные галочки
        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections()
                : java.util.Collections.emptyMap();

        if (index >= 0) {
            building.getFloors().remove(index);
            floorListModel.remove(index);
        }

        // Перерисовываем без force и без авто-правил, затем возвращаем снимок
        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        if (rt != null) rt.restoreSelections(snap);

        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);
        updateNoiseTabNow();
    }
    private void addSpace(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) {
            showMessage("Выберите этаж!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        AddSpaceDialog dialog = new AddSpaceDialog((JFrame) SwingUtilities.getWindowAncestor(this), selectedFloor.getType());
        Space space = null;
        if (dialog.showDialog()) {
            space = new Space();
            space.setIdentifier(dialog.getSpaceIdentifier());
            // 1) выбор пользователя
            space.setType(dialog.getSpaceType());

            // 2) запрет OUTDOOR на смешанном этаже
            if (selectedFloor.getType() == Floor.FloorType.MIXED && space.getType() == Space.SpaceType.OUTDOOR) {
                space.setType(Space.SpaceType.APARTMENT);
            }
            // 3) на «улице» принудительно OUTDOOR
            if (selectedFloor.getType() == Floor.FloorType.STREET) {
                space.setType(Space.SpaceType.OUTDOOR);
            }

            // позиция в КОНЕЦ текущего этажа
            int maxPos = selectedFloor.getSpaces().stream()
                    .mapToInt(Space::getPosition)
                    .max().orElse(-1);
            space.setPosition(maxPos + 1);

            selectedFloor.addSpace(space);

            // офисам — МК по умолчанию (кроме санузлов)
            applyMicroDefaultsForOfficeSpace(space);

            // ПЕРЕСОБРАТЬ список с учётом сортировки и фильтров
            updateSpaceList();
            if (spaceListModel.getSize() > 0) {
                spaceList.setSelectedValue(space, true);
            }
        }

        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections() : java.util.Collections.emptyMap();

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        if (rt != null) {
            rt.restoreSelections(snap); // вернули все старые

            // ВАЖНО: если на этом этаже уже есть помещения с галочками — НИЧЕГО не проставляем в новом
            Floor f = floorList.getSelectedValue();
            if (f != null && !rt.hasAnySelectedOnFloor(f)) {
                rt.applyDefaultsForSpace(space); // на этаже ещё пусто — можно поставить дефолты в новое помещение
            }

            // Выделим новое помещение в таблице радиации
            if (f != null && space != null) {
                int newIndex = f.getSpaces().indexOf(space);
                if (newIndex >= 0) rt.selectSpaceByIndex(newIndex);
            }
        }
        updateNoiseTabNow();
    }

    private void editSpace(ActionEvent e) {
        Floor floor = floorList.getSelectedValue();
        int index = spaceList.getSelectedIndex();

        if (floor == null || index < 0) {
            showMessage("Выберите помещение для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Space space = spaceListModel.get(index);
        AddSpaceDialog dialog = new AddSpaceDialog((JFrame) SwingUtilities.getWindowAncestor(this), floor.getType());
        dialog.setSpaceIdentifier(space.getIdentifier());
        dialog.setSpaceType(space.getType());

        if (dialog.showDialog()) {
            space.setIdentifier(dialog.getSpaceIdentifier());
            space.setType(dialog.getSpaceType());

            // запрет OUTDOOR на смешанном этаже
            if (floor.getType() == Floor.FloorType.MIXED && space.getType() == Space.SpaceType.OUTDOOR) {
                space.setType(Space.SpaceType.APARTMENT);
            }
            // на «улице» — только OUTDOOR
            if (floor.getType() == Floor.FloorType.STREET) {
                space.setType(Space.SpaceType.OUTDOOR);
            }

            applyMicroDefaultsForOfficeSpace(space);
            spaceListModel.set(index, space);
            updateRoomList();
        }

        // Радиация — трогаем только UI, старые галочки сохраняем
        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections()
                : java.util.Collections.emptyMap();

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        if (rt != null) rt.restoreSelections(snap);

        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);
        updateNoiseTabNow();
    }

    private void removeSpace(ActionEvent e) {
        // Сохраняем радиационные галочки заранее
        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections()
                : java.util.Collections.emptyMap();

        Floor floor = floorList.getSelectedValue();
        int index = spaceList.getSelectedIndex();
        if (floor != null && index >= 0) {
            floor.getSpaces().remove(index);
            spaceListModel.remove(index);
        }

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        if (rt != null) rt.restoreSelections(snap);

        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        updateNoiseTabNow();
    }

    private void addRoom(ActionEvent e) {
        Space selectedSpace = spaceList.getSelectedValue();
        if (selectedSpace == null) {
            showMessage("Выберите помещение!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }
        int roomsBefore = selectedSpace.getRooms().size(); // было ли помещение пустым

        // ===== Снимок значений вкладки «Осв улица», чтобы их не потерять после refresh =====
        StreetLightingTab streetBefore = getStreetLightingTab();
        java.util.Map<String, Double[]> streetSnapshot = java.util.Collections.emptyMap();
        if (streetBefore != null) {
            try {
                streetSnapshot = streetBefore.snapshotValuesByKey();
            } catch (Throwable ignore) {
                // если что-то пойдёт не так — просто не восстановим значения
            }
        }

        // текущий этаж (нужен, чтобы понять «улица»)
        Floor currentFloor = (floorList != null) ? floorList.getSelectedValue() : null;
        boolean isStreetFloor = (currentFloor != null && currentFloor.getType() == Floor.FloorType.STREET);

        if (isApartmentSpace(selectedSpace)) {
            // квартиры — AddRoomDialog
            java.util.List<String> suggestions = collectPopularApartmentRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddRoomDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddRoomDialog(parent, suggestions, "", false);
            if (dlg.showDialog()) {
                java.util.List<String> names = dlg.getNamesToAddList();
                for (String name : names) {
                    if (name == null || name.isBlank()) continue;
                    Room room = new Room();
                    room.setName(name.trim());
                    room.setSelected(false);
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    selectedSpace.addRoom(room);
                    roomListModel.addElement(room);
                }
            }
        } else if (isOfficeSpace(selectedSpace)) {
            // офисы — AddOfficeRoomsDialog
            java.util.List<String> suggestions = collectPopularOfficeRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddOfficeRoomsDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddOfficeRoomsDialog(parent, suggestions, "", false);
            if (dlg.showDialog()) {
                java.util.List<String> names = dlg.getNamesToAddList();
                for (String name : names) {
                    if (name == null || name.isBlank()) continue;
                    Room room = new Room();
                    room.setName(name.trim());
                    room.setSelected(false);
                    room.setMicroclimateSelected(!looksLikeSanitary(room.getName()));
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    selectedSpace.addRoom(room);
                    roomListModel.addElement(room);
                }
            }
        } else if (isPublicSpace(selectedSpace)) {
            // общественные — AddPublicRoomsDialog
            List<String> suggestions = collectPopularPublicRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddPublicRoomsDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddPublicRoomsDialog(parent, suggestions, "", false);
            if (dlg.showDialog()) {
                List<String> names = dlg.getNamesToAddList();
                for (String name : names) {
                    if (name == null || name.isBlank()) continue;
                    Room room = new Room();
                    room.setName(name.trim());
                    room.setSelected(false);
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    selectedSpace.addRoom(room);
                    roomListModel.addElement(room);
                }
            }
        } else if (isStreetFloor || isStreetSpace(selectedSpace)) {
            // УЛИЦА — даем готовые подсказки (без цифр)
            List<String> suggestions = collectStreetRoomSuggestions();
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddRoomDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddRoomDialog(parent, suggestions, "", false);
            if (dlg.showDialog()) {
                List<String> names = dlg.getNamesToAddList();
                for (String name : names) {
                    if (name == null || name.isBlank()) continue;
                    Room room = new Room();
                    room.setName(name.trim());
                    room.setSelected(false);
                    // для улицы — логика стен оставляем как общую (0 для санузлов, 1 иначе)
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    selectedSpace.addRoom(room);
                    roomListModel.addElement(room);
                }
            }
        } else {
            String name = showInputDialog("Добавление комнаты", "Название комнаты:", "");
            if (name != null && !name.isBlank()) {
                Room room = new Room();
                room.setName(name.trim());
                room.setSelected(false);
                try {
                    int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                    room.setExternalWallsCount(walls);
                } catch (Throwable ignore) {}
                selectedSpace.addRoom(room);
                roomListModel.addElement(room);
            }
        }

        // Радиация — безопасная перерисовка
        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections()
                : java.util.Collections.emptyMap();

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        if (rt != null) {
            rt.restoreSelections(snap);
            // Если помещение было пустым — дефолты радиации ТОЛЬКО для него
            if (roomsBefore == 0) {
                Floor f = (floorList != null) ? floorList.getSelectedValue() : null;
                if (f != null && !rt.hasAnySelectedOnFloor(f)) {
                    rt.applyDefaultsForSpace(selectedSpace);
                }
            }
        }

        // Эти вызовы перерисуют и вкладку «Осв улица»
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        // ===== Вернём значения «Осв улица» после refresh =====
        if (streetSnapshot != null && !streetSnapshot.isEmpty()) {
            StreetLightingTab streetAfter = getStreetLightingTab();
            if (streetAfter != null) {
                try {
                    streetAfter.applyValuesByKey(streetSnapshot);
                } catch (Throwable ignore) {
                    // тихо игнорируем сбой восстановления
                }
            }
        }
        updateNoiseTabNow();
    }


    private void editRoom(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space == null || index < 0) {
            showMessage("Выберите комнату для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Room room = roomListModel.get(index);

        if (isApartmentSpace(space)) {
            List<String> suggestions = collectPopularApartmentRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddRoomDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddRoomDialog(parent, suggestions, room.getName(), true);
            if (dlg.showDialog()) {
                String newName = dlg.getNameToAdd();
                if (newName != null && !newName.isBlank()) {
                    room.setName(newName.trim());
                    roomListModel.set(index, room);
                }
            }
        } else if (isOfficeSpace(space)) {
            List<String> suggestions = collectPopularOfficeRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddOfficeRoomsDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddOfficeRoomsDialog(parent, suggestions, room.getName(), true);
            if (dlg.showDialog()) {
                String newName = dlg.getNameToAdd();
                if (newName != null && !newName.isBlank()) {
                    room.setName(newName.trim());
                    room.setMicroclimateSelected(!looksLikeSanitary(room.getName())); // офис: освежаем МК-галочку
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    roomListModel.set(index, room);
                }
            }
        } else if (isPublicSpace(space)) {
            List<String> suggestions = collectPopularPublicRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddPublicRoomsDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddPublicRoomsDialog(parent, suggestions, room.getName(), true);
            if (dlg.showDialog()) {
                String newName = dlg.getNameToAdd();
                if (newName != null && !newName.isBlank()) {
                    room.setName(newName.trim());
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    roomListModel.set(index, room);
                }
            }
        } else {
            String newName = showInputDialog("Редактирование комнаты", "Новое название комнаты:", room.getName());
            if (newName != null && !newName.isBlank()) {
                room.setName(newName.trim());
                roomListModel.set(index, room);
            }
        }

        // Радиация — безопасная перерисовка
        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections()
                : java.util.Collections.emptyMap();

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        if (rt != null) rt.restoreSelections(snap);

        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);
        updateNoiseTabNow();
    }

    private void removeRoom(ActionEvent e) {
        // Сохранить текущие галочки радиации
        RadiationTab rt = getRadiationTab();
        Map<String, Boolean> snap = (rt != null) ? rt.saveSelections()
                : java.util.Collections.emptyMap();

        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space != null && index >= 0) {
            space.getRooms().remove(index);
            roomListModel.remove(index);
        }

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        if (rt != null) rt.restoreSelections(snap);

        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        updateNoiseTabNow();
    }

    private void updateRoomList() {
        roomListModel.clear();
        Space selectedSpace = spaceList.getSelectedValue();

        if (selectedSpace != null) {
            for (Room room : selectedSpace.getRooms()) {
                roomListModel.addElement(room);
            }
        }
    }

    private void refreshAllLists() {
        // Секции
        refreshSectionListModel();
        if (!sectionListModel.isEmpty() && sectionList.getSelectedIndex() < 0) {
            sectionList.setSelectedIndex(0);
        }

        // Этажи выбранной секции
        refreshFloorListForSelectedSection();
        updateNoiseTabNow();

        // Помещения/комнаты для актуального выбора
        updateSpaceList();
        updateRoomList();
    }

    private void updateLightingTab(Building building, boolean autoApplyDefaults) {
        // КЕО — как было
        LightingTab keo = getLightingTab();
        if (keo != null) {
            keo.setBuilding(building, autoApplyDefaults);
            keo.refreshData();
        }
        // НОВОЕ: отдельная вкладка «Освещение» со своей картой галочек
        ArtificialLightingTab alt = getArtificialLightingTab();
        if (alt != null) {
            alt.setBuilding(building, autoApplyDefaults); // при true – авто-галочки для офис/общественных
            alt.refreshData();
        }
        // Осв улица — ПОДТЯГИВАЕМ АКТУАЛЬНЫЙ BUILDING
        ru.citlab24.protokol.tabs.modules.lighting.StreetLightingTab street = getStreetLightingTab();
        if (street != null) {
            street.setBuilding(building);  // без авто-дефолтов, просто перечитать названия
            street.refreshData();
        }
    }

    private void updateVentilationTab(Building building) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            Arrays.stream(((MainFrame) mainFrame).getTabbedPane().getComponents())
                    .filter(tab -> tab instanceof VentilationTab)
                    .findFirst()
                    .ifPresent(tab -> {
                        ((VentilationTab) tab).setBuilding(building);
                        ((VentilationTab) tab).refreshData();
                    });
        }
    }

    private String showInputDialog(String title, String message, String initialValue) {
        JTextField field = new JTextField(initialValue);

        JPanel panel = new JPanel(new BorderLayout(8, 0));
        panel.add(new JLabel(message), BorderLayout.WEST);
        panel.add(field, BorderLayout.CENTER);

        // Создаём диалог вручную, чтобы управлять фокусом и клавишами
        final JOptionPane pane = new JOptionPane(
                panel,
                JOptionPane.PLAIN_MESSAGE,
                JOptionPane.OK_CANCEL_OPTION
        );
        final JDialog dialog = pane.createDialog(this, title);
        dialog.setModal(true);
        dialog.setResizable(false);

        // Enter = OK
        field.addActionListener(e -> {
            pane.setValue(JOptionPane.OK_OPTION);
            dialog.dispose();
        });

        // Esc = Cancel
        dialog.getRootPane().registerKeyboardAction(
                e -> {
                    pane.setValue(JOptionPane.CANCEL_OPTION);
                    dialog.dispose();
                },
                KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0),
                JComponent.WHEN_IN_FOCUSED_WINDOW
        );

        // Автофокус и выделение текста сразу после показа
        dialog.addWindowListener(new WindowAdapter() {
            @Override public void windowOpened(WindowEvent e) {
                field.requestFocusInWindow();
                field.selectAll();
            }
        });

        dialog.setVisible(true);

        Object value = pane.getValue();
        boolean ok = (value != null) && Integer.valueOf(JOptionPane.OK_OPTION).equals(value);
        return ok ? field.getText() : null;
    }

    private void showMessage(String message, String title, int messageType) {
        JOptionPane.showMessageDialog(this, message, title, messageType);
    }

    private void handleError(String message, String title) {
        // Логирование ошибки
        System.err.println(title + ": " + message);
        showMessage(message, title, JOptionPane.ERROR_MESSAGE);
    }

    @Override
    public void addNotify() {
        super.addNotify();
        if (this.building == null) this.building = new Building();
        updateRadiationTab(this.building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(building, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(building, false);

        if (floorList != null && !floorListModel.isEmpty()) {
            floorList.setSelectedIndex(0);
        }
        if (floorList != null) {
            floorList.addListSelectionListener(evt -> {
                if (!evt.getValueIsAdjusting()) {
                    // Никаких автопереключений тумблеров — только обновляем список помещений
                    updateSpaceList();
                }
            });
        }

        if (spaceList != null) {
            spaceList.addListSelectionListener(evt -> {
                if (!evt.getValueIsAdjusting()) updateRoomList();
            });
        }
    }

    private java.util.List<String> collectPopularApartmentRoomNames(Building b, int limit) {
        return ops.collectPopularApartmentRoomNames(limit);
    }

    // Возвращает true, если помещение — «квартира».
// Не опираемся на конкретные enum-константы проекта, работаем по имени типа и по идентификатору.
    private boolean isApartmentSpace(Space s) { return ops.isApartmentSpace(s); }


    // Офисное помещение?
    private boolean isOfficeSpace(Space s) { return ops.isOfficeSpace(s); }

    // Общественное помещение?
    private boolean isPublicSpace(Space s) { return ops.isPublicSpace(s); }

    // Топ общественных названий по зданию
    private java.util.List<String> collectPopularPublicRoomNames(Building b, int limit) {
        return ops.collectPopularPublicRoomNames(limit);
    }

    // Сбор топа офисных названий по зданию (по "базе", без номеров)
    private java.util.List<String> collectPopularOfficeRoomNames(Building b, int limit) {
        return ops.collectPopularOfficeRoomNames(limit);
    }

    // КОПИЯ ДЛЯ СОХРАНЕНИЯ: сохраняем id и selected
    private Floor createFloorCopyPreserve(Floor original) { return ops.createFloorCopyPreserve(original); }


    private Space createSpaceCopyPreserve(Space original) { return ops.createSpaceCopyPreserve(original); }

    private LightingTab getLightingTab() {
        Window wnd = SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof MainFrame) {
            JTabbedPane tabs = ((MainFrame) wnd).getTabbedPane();
            for (Component c : tabs.getComponents()) {
                if (c instanceof LightingTab) return (LightingTab) c;
            }
        }
        return null;
    }
    private ArtificialLightingTab getArtificialLightingTab() {
        Window wnd = SwingUtilities.getWindowAncestor(this);
        return (wnd instanceof MainFrame) ? ((MainFrame) wnd).getArtificialLightingTab() : null;
    }

    private StreetLightingTab getStreetLightingTab() {
        java.awt.Window wnd = javax.swing.SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof ru.citlab24.protokol.MainFrame) {
            return ((ru.citlab24.protokol.MainFrame) wnd).getStreetLightingTab();
        }
        return null;
    }


    private void updateMicroclimateTab(Building building, boolean autoApplyDefaults) {
        Window wnd = SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof MainFrame) {
            JTabbedPane tabs = ((MainFrame) wnd).getTabbedPane();
            for (Component c : tabs.getComponents()) {
                if (c instanceof MicroclimateTab) {
                    ((MicroclimateTab) c).display(building, autoApplyDefaults);
                    break;
                }
            }
        }
    }

    public Building getCurrentBuilding() {
        return building;
    }
    private MicroclimateTab getMicroclimateTab() {
        java.awt.Window wnd = javax.swing.SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof ru.citlab24.protokol.MainFrame) {
            javax.swing.JTabbedPane tabs = ((ru.citlab24.protokol.MainFrame) wnd).getTabbedPane();
            for (java.awt.Component c : tabs.getComponents()) {
                if (c instanceof ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateTab) {
                    return (ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateTab) c;
                }
            }
        }
        return null;
    }
    private static boolean looksLikeSanitary(String name) {
        return BuildingModelOps.looksLikeSanitary(name);
    }

    // ===== МИКРОКЛИМАТ: сохранение/восстановление по ключу "этаж|помещение|комната" =====
    private Map<String, Boolean> saveMicroclimateSelections() { return ops.saveMicroclimateSelections(); }

    private void restoreMicroclimateSelections(Map<String, Boolean> saved) { ops.restoreMicroclimateSelections(saved); }

    /** ===== КЕО (естественное освещение): снимок и восстановление по ключу "этаж|помещение|комната" ===== */

    /** КЕО: сделать снимок состояний чекбоксов (Room.isSelected) по всему зданию. */
    private Map<String, Boolean> saveKeoSelections() {
        Map<String, Boolean> res = new HashMap<>();
        if (building == null) return res;

        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                for (Room r : s.getRooms()) {
                    String key = f.getNumber() + "|" + s.getIdentifier() + "|" + r.getName();
                    res.put(key, r.isSelected());
                }
            }
        }
        return res;
    }

    /** КЕО: восстановить состояния чекбоксов в ПЕРЕДАННОМ building по снимку. */
    private static void restoreKeoSelections(Building target, Map<String, Boolean> saved) {
        if (target == null || saved == null || saved.isEmpty()) return;

        for (Floor f : target.getFloors()) {
            for (Space s : f.getSpaces()) {
                for (Room r : s.getRooms()) {
                    String key = f.getNumber() + "|" + s.getIdentifier() + "|" + r.getName();
                    Boolean v = saved.get(key);
                    if (v != null) {
                        r.setSelected(v);
                    }
                }
            }
        }
    }

    /** Микроклимат: для офисов (по типу помещения ИЛИ по типу этажа) ставим всем, кроме санузлов */
    private void applyMicroDefaultsForOfficeSpace(Space s) { ops.applyMicroDefaultsForOfficeSpace(s); }

    /** true, если помещение проходит текущие тумблеры-фильтры. */
    private boolean isSpaceVisibleByFilter(Space s) {
        boolean showA = (filterApartmentBtn == null) || filterApartmentBtn.isSelected();
        boolean showO = (filterOfficeBtn    == null) || filterOfficeBtn.isSelected();
        boolean showP = (filterPublicBtn    == null) || filterPublicBtn.isSelected();
        boolean showS = (filterStreetBtn    == null) || filterStreetBtn.isSelected();

        boolean isA = isApartmentSpace(s);
        boolean isO = isOfficeSpace(s);
        boolean isP = isPublicSpace(s);
        boolean isS = (s != null && s.getType() == Space.SpaceType.OUTDOOR);

        if (isA && showA) return true;
        if (isO && showO) return true;
        if (isP && showP) return true;
        if (isS && showS) return true;

        // Нераспознанный тип показываем только если включены все тумблеры
        if (!isA && !isO && !isP && !isS) return (showA && showO && showP && showS);

        return false;
    }
    // На «улице» показываем только OUTDOOR и только если включён тумблер «Улица»
    private boolean isSpaceVisibleByFilterOnFloor(Space s, Floor f) {
        if (f != null && f.getType() == Floor.FloorType.STREET) {
            boolean showS = (filterStreetBtn == null) || filterStreetBtn.isSelected();
            return showS && s.getType() == Space.SpaceType.OUTDOOR;
        }
        return isSpaceVisibleByFilter(s);
    }
    /** Подсказки для «улицы» (без чисел). */
    private java.util.List<String> collectStreetRoomSuggestions() {
        return java.util.List.of(
                "дорога",
                "пешеходная дорожка у входа в здание",
                "аллея",
                "пожарные проезды",
                "тротуары-подъезды",
                "автостоянка",
                "хозяйственная площадка",
                "площадка при мусоросборниках",
                "прогулочная дорожка",
                "физкультурные площадки",
                "площадки для игр",
                "основной вход в здание",
                "запасной вход в здание",
                "технический вход в здание"
        );
    }

    /** true, если помещение уличное (OUTDOOR). */
    private boolean isStreetSpace(Space s) {
        return s != null && s.getType() == Space.SpaceType.OUTDOOR;
    }
    private NoiseTab getNoiseTab() {
        java.awt.Window wnd = javax.swing.SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof ru.citlab24.protokol.MainFrame) {
            javax.swing.JTabbedPane tabs = ((ru.citlab24.protokol.MainFrame) wnd).getTabbedPane();
            for (java.awt.Component c : tabs.getComponents()) {
                if (c instanceof NoiseTab) {
                    return (NoiseTab) c;
                }
            }
        }
        return null;
    }

    /** НОВОЕ: мгновенно обновляет вкладку «Шумы» от текущей модели building, без сохранения в БД. */
    private void updateNoiseTabNow() {
        try {
            NoiseTab noise = getNoiseTab();
            if (noise == null) return;

            // Сохраним текущие галочки/источники и восстановим их после перестройки
            noise.updateRoomSelectionStates();
            java.util.Map<String, ru.citlab24.protokol.db.DatabaseManager.NoiseValue> snap = noise.saveSelectionsByKey();

            noise.setBuilding(building);      // подсовываем актуальную модель из памяти
            noise.applySelectionsByKey(snap); // вернули галочки по ключу
            noise.refreshData();              // перерисовали UI
        } catch (Throwable ignore) {
            // тихо игнорируем, чтобы ни одна операция добавления/копирования не падала
        }
    }


}