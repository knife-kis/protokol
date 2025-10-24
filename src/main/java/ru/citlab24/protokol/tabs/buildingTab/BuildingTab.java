package ru.citlab24.protokol.tabs.buildingTab;

import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.dialogs.*;
import ru.citlab24.protokol.tabs.modules.lighting.LightingTab;
import ru.citlab24.protokol.tabs.modules.med.RadiationTab;
import ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateTab;
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
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
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
    private static final Color FLOOR_PANEL_COLOR = new Color(0, 115, 200);
    private static final Color SPACE_PANEL_COLOR = new Color(76, 175, 80);
    private static final Color ROOM_PANEL_COLOR = new Color(156, 39, 176);
    private static final Font HEADER_FONT =
            UIManager.getFont("Label.font").deriveFont(Font.PLAIN, 15f);
    private static final Dimension BUTTON_PANEL_SIZE = new Dimension(5, 5);

    private Building building;
    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final DefaultListModel<Space> spaceListModel = new DefaultListModel<>();
    private final DefaultListModel<Room> roomListModel = new DefaultListModel<>();
    private final DefaultListModel<Section> sectionListModel = new DefaultListModel<>();
    private JList<Section> sectionList;

    private JList<Floor> floorList;
    private JList<Space> spaceList;
    private JList<Room> roomList;
    private JTextField projectNameField;

    public BuildingTab(Building building) {
        // 1) сохраняем в поле, 2) создаём дефолт, если пришёл null
        this.building = (building != null) ? building : new Building();

        // гарантируем хотя бы одну секцию
        if (this.building.getSections().isEmpty()) {
            this.building.addSection(new Section("Секция 1", 0));
        }

        initComponents();
    }
    // === НОРМАЛИЗАЦИЯ БАЗОВОГО НАЗВАНИЯ КОМНАТЫ (без площадей/осей/скобок) ===
    private static final java.util.regex.Pattern TAIL_PLO =
            java.util.regex.Pattern.compile(
                    "\\s*[,;]?\\s*площад[ьяиуе][^\\d]*\\d+[\\d.,]*\\s*(?:кв\\.?\\s*м|м\\s*²|м2|м\\^2)\\.?\\s*$",
                    java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE);

    private static final java.util.regex.Pattern TAIL_AXES =
            java.util.regex.Pattern.compile(
                    "\\s*[,;]?\\s*в\\s+осях\\s+[\\p{L}\\p{N}]+(?:\\s*[-—–]\\s*[\\p{L}\\p{N}]+)?\\s*$",
                    java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE);

    // Удаляем хвостовые скобки вида "(...)" в конце
    private static final java.util.regex.Pattern TAIL_PAREN =
            java.util.regex.Pattern.compile("\\s*\\((?:[^)(]+|\\([^)(]*\\))*\\)\\s*$");

    private static String normalizeRoomBaseName(String name) {
        if (name == null) return "";
        String s = name.replace('\u00A0',' ').trim();      // NBSP → пробел
        s = s.replaceAll("\\s+", " ");

        // Сначала убираем финальные скобки "(площадью 15 кв. м)" / "(в осях 1-2)" и т.п.
        boolean changed;
        do {
            changed = false;
            java.util.regex.Matcher mp = TAIL_PAREN.matcher(s);
            if (mp.find()) { s = s.substring(0, mp.start()).trim(); changed = true; }
        } while (changed);

        // Затем убираем явные хвосты "в осях …" и "площадью …"
        s = TAIL_AXES.matcher(s).replaceAll("");
        s = TAIL_PLO.matcher(s).replaceAll("");

        // Чистим завершающие знаки/пробелы
        s = s.replaceAll("[\\s\\.,;:—–-]+$", "").trim();

        return s;
    }


    private void setRussianLocale() {
        Locale.setDefault(new Locale("ru", "RU"));
        UIManager.put("OptionPane.yesButtonText", "Да");
        UIManager.put("OptionPane.noButtonText", "Нет");
        UIManager.put("OptionPane.cancelButtonText", "Отмена");

        // Оптимизированная установка локали для UIManager
        UIManager.getLookAndFeelDefaults().keySet().stream()
                .filter(key -> key instanceof String && ((String) key).endsWith(".locale"))
                .forEach(key -> UIManager.getLookAndFeelDefaults().put(key, Locale.getDefault()));
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
        panel.add(createListPanel("Помещения на этаже", SPACE_PANEL_COLOR, spaceListModel,
                new SpaceListRenderer(), this::createSpaceButtons));
        panel.add(createListPanel("Комнаты в помещении", ROOM_PANEL_COLOR, roomListModel,
                new RoomListRenderer(), this::createRoomButtons));
        return panel;
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
            // (перестановку этажей внутри секции мы вешаем в createSectionsAndFloorsPanel, тут не нужно)
        } else if ("Помещения на этаже".equals(title)) {
            spaceList = (JList<Space>) list;
            // <<< ВОТ ЭТИ ТРИ СТРОКИ >>>
            spaceList.setDragEnabled(true);
            spaceList.setDropMode(DropMode.INSERT);
            spaceList.setTransferHandler(spaceReorderHandler);
        } else if ("Комнаты в помещении".equals(title)) {
            roomList = (JList<Room>) list;
            // <<< И ВОТ ЭТИ ТРИ >>>
            roomList.setDragEnabled(true);
            roomList.setDropMode(DropMode.INSERT);
            roomList.setTransferHandler(roomReorderHandler);
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

    private JPanel createActionButtons() {
        return createButtonPanel(
                createStyledButton("Загрузить проект", FontAwesomeSolid.FOLDER_OPEN, new Color(33, 150, 243), this::loadProject),
                createStyledButton("Сохранить проект", FontAwesomeSolid.SAVE, new Color(0, 115, 200), this::saveProject),
                createStyledButton("Рассчитать показатели", FontAwesomeSolid.CALCULATOR, new Color(103, 58, 183), this::calculateMetrics),
                // НОВОЕ: кнопка-отчёт
                createStyledButton("Сводка квартир", FontAwesomeSolid.TABLE, new Color(96, 125, 139), this::showApartmentSummary)
        );
    }

    private void copySection(ActionEvent e) {
        if (sectionList == null || sectionList.isSelectionEmpty()) {
            showMessage("Выберите секцию для копирования", "Информация", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        int srcIdx = sectionList.getSelectedIndex();
        Section src = sectionListModel.get(srcIdx);

        // спросим имя новой секции
        String suggested = generateUniqueSectionName(src.getName());
        String newName = showInputDialog("Копировать секцию", "Название новой секции:", suggested);
        if (newName == null || newName.trim().isEmpty()) return;

        // создаём новую секцию в конце списка
        int newPosition = building.getSections().size();
        Section dst = new Section();
        dst.setName(newName.trim());
        dst.setPosition(newPosition);
        building.addSection(dst);

        // индекс новой секции
        int dstIdx = newPosition;

        // Копируем все этажи исходной секции
        java.util.List<Floor> toAdd = new java.util.ArrayList<>();
        for (Floor f : building.getFloors()) {
            if (f.getSectionIndex() == srcIdx) {
                Floor fCopy = createFloorCopy(f);               // уже умеет глубоко копировать помещения/комнаты
                fCopy.setSectionIndex(dstIdx);                  // привязываем к новой секции
                toAdd.add(fCopy);
            }
        }
        // добавляем скопированные этажи в здание
        for (Floor fCopy : toAdd) {
            building.addFloor(fCopy);
        }

        // UI: показать новую секцию и её этажи
        refreshSectionListModel();
        sectionList.setSelectedIndex(dstIdx);
        refreshFloorListForSelectedSection();
        updateSpaceList();
        updateRoomList();

        // Обновим вкладки расчётов
        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateVentilationTab(building);

// ВАЖНО: обновляем Освещение и запускаем авто-проставление
// → на НОВОЙ секции отметится только самый нижний жилой/совмещённый этаж,
// существующие галочки не трогаются.
        updateLightingTab(building, /*autoApplyDefaults=*/true);

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

// начальная выборка
        if (!sectionListModel.isEmpty()) sectionList.setSelectedIndex(0);
        refreshFloorListForSelectedSection();

// DnD: секции ↔ секции (перестановка) И перенос этажей на секцию
        sectionList.setDragEnabled(true);
        sectionList.setDropMode(DropMode.INSERT);
        sectionList.setTransferHandler(sectionReorderHandler);

// DnD: перестановка этажей внутри секции
        floorList.setDragEnabled(true);
        floorList.setDropMode(DropMode.INSERT);
        floorList.setTransferHandler(floorReorderHandler);



// Этажи: перестановка внутри секции
        floorList.setDragEnabled(true);
        floorList.setDropMode(DropMode.INSERT);
        floorList.setTransferHandler(floorReorderHandler);


        // слушатели
        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) refreshFloorListForSelectedSection();
        });

        // раскладка
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

    // Перестановка секций МЕЖДУ собой + приём "этажа" (перенос этажа на секцию)
    // Перестановка секций МЕЖДУ собой (ПЕРЕНОС ЭТАЖЕЙ НА СЕКЦИЮ ЗАПРЕЩЁН)
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
        int secIdx = Math.max(0, sectionList.getSelectedIndex());
        java.util.List<Floor> list = new java.util.ArrayList<>();
        for (Floor f : building.getFloors())
            if (f.getSectionIndex() == secIdx) list.add(f);
        list.sort(java.util.Comparator.comparingInt(Floor::getPosition));
        floorListModel.clear();
        list.forEach(floorListModel::addElement);

        if (!floorListModel.isEmpty()) floorList.setSelectedIndex(0);
    }


    private JButton createStyledButton(String text, FontAwesomeSolid icon, Color bgColor, ActionListener action) {
        JButton btn = new JButton(text, FontIcon.of(icon, 16, Color.WHITE));
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
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

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

    private void loadSelectedProject(Building selectedProject) throws SQLException {
        Building loadedBuilding = DatabaseManager.loadBuilding(selectedProject.getId());
        this.building = loadedBuilding;
        projectNameField.setText(loadedBuilding.getName());
        refreshAllLists();
        updateVentilationTab(loadedBuilding);
        updateRadiationTab(loadedBuilding, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(loadedBuilding, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(loadedBuilding, /*autoApplyDefaults=*/false);
        showMessage("Проект '" + loadedBuilding.getName() + "' успешно загружен", "Загрузка", JOptionPane.INFORMATION_MESSAGE);
    }

    private void saveProject(ActionEvent e) {
        logger.info("BuildingTab.saveProject() - Начало сохранения проекта");

        // 0) Финализируем активное редактирование таблиц (чтобы текущее значение попало в модель)
        try {
            java.awt.KeyboardFocusManager kfm = java.awt.KeyboardFocusManager.getCurrentKeyboardFocusManager();
            java.awt.Component fo = (kfm != null) ? kfm.getFocusOwner() : null;
            JTable editingTable = (fo == null) ? null
                    : (JTable) javax.swing.SwingUtilities.getAncestorOfClass(JTable.class, fo);
            if (editingTable != null && editingTable.isEditing()) {
                try { editingTable.getCellEditor().stopCellEditing(); } catch (Exception ignore) {}
            }
        } catch (Exception ignore) {}

        // 1) Синхронизируем UI → модель (радиация)
        RadiationTab radiationTab = getRadiationTab();
        if (radiationTab != null) {
            radiationTab.updateRoomSelectionStates(); // Сохраняем ручные изменения
        }

        // 1.1) Пробрасываем выбранность из «Освещения»
        LightingTab lightingTab = getLightingTab();
        if (lightingTab != null) {
            lightingTab.updateRoomSelectionStates();
        }

        // 1.2) Пробрасываем выбранность из «Микроклимата»
        MicroclimateTab microTab = getMicroclimateTab();
        if (microTab != null) {
            microTab.updateRoomSelectionStates();
        }

        // 2) Генерация имени проекта
        String baseName = projectNameField.getText().trim();
        if (baseName.isEmpty()) {
            showMessage("Введите название проекта!", "Ошибка", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 3) Создаем копию проекта с сохранением состояний
        Building newProject = createBuildingCopy();
        newProject.setName(generateProjectVersionName(baseName));

        // 4) Сохраняем в БД
        try {
            DatabaseManager.saveBuilding(newProject);
            this.building = newProject;
            projectNameField.setText(extractBaseName(newProject.getName()));
        } catch (SQLException ex) {
            handleError("Ошибка сохранения: " + ex.getMessage(), "Ошибка");
            return;
        }

        // 5) Переинициализируем вкладки без авто-проставления
        updateRadiationTab(newProject, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(newProject, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(newProject, /*autoApplyDefaults=*/false);

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
            updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

        }
    };

    private final TransferHandler roomReorderHandler = new ReorderHandler<Room>() {
        @Override protected void afterReorder(DefaultListModel<Room> model) {
            for (int i = 0; i < model.size(); i++) model.get(i).setPosition(i);
            updateLightingTab(building, /*autoApplyDefaults=*/true);
            updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

        }
    };

    private Floor createFloorCopy(Floor originalFloor) {
        Floor floorCopy = new Floor();
        floorCopy.setNumber(originalFloor.getNumber());
        floorCopy.setType(originalFloor.getType());
        floorCopy.setName(originalFloor.getName());
        floorCopy.setSectionIndex(originalFloor.getSectionIndex()); // ВАЖНО
        floorCopy.setPosition(originalFloor.getPosition());          // ← сохраняем порядок в секции

        for (Space origSpace : originalFloor.getSpaces()) {
            Space spaceCopy = createSpaceCopyWithNewIds(origSpace);
            floorCopy.addSpace(spaceCopy);
        }
        return floorCopy;

    }


    public void refreshFloor(Floor floor) {
        int index = building.getFloors().indexOf(floor);
        if (index >= 0) {
            floorListModel.set(index, floor);

            if (floorList.getSelectedIndex() == index) {
                updateSpaceList();
                spaceList.setSelectedIndex(0);
            }
        }
    }

    private Room createRoomCopy(Room originalRoom) {
        Room roomCopy = new Room();
        roomCopy.setId(generateUniqueRoomId());
        roomCopy.setName(originalRoom.getName());
        roomCopy.setVolume(originalRoom.getVolume());
        roomCopy.setVentilationChannels(originalRoom.getVentilationChannels());
        roomCopy.setVentilationSectionArea(originalRoom.getVentilationSectionArea());
        roomCopy.setSelected(false);
        roomCopy.setOriginalRoomId(originalRoom.getId()); // Сохраняем ссылку на оригинал
        return roomCopy;
    }

    private int generateUniqueRoomId() {
        return UUID.randomUUID().hashCode();
    }

    private void calculateMetrics(ActionEvent e) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            updateVentilationTab(building);
            updateRadiationTab(building, true);
            ; // Добавлено обновление RadiationTab
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

        // Сохраняем состояния чекбоксов ДО копирования
        Map<String, Boolean> savedSelections = saveRadiationSelections();

        // Создаем копию помещения с НОВЫМИ ID комнат
        Space copiedSpace = createSpaceCopyWithNewIds(selectedSpace);
        copiedSpace.setIdentifier(generateNextSpaceId(selectedFloor, selectedSpace.getIdentifier()));

        // Добавляем новое помещение
        selectedFloor.addSpace(copiedSpace);
        spaceListModel.addElement(copiedSpace);
        spaceList.setSelectedValue(copiedSpace, true);

        // Обновляем вкладку с новым зданием
        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

        // Восстанавливаем состояния ТОЛЬКО для исходных комнат
        restoreRadiationSelections(savedSelections);

        // Явно выделяем новое помещение в RadiationTab
        RadiationTab radiationTab = getRadiationTab();
        if (radiationTab != null) {
            // Находим индекс нового помещения
            int newIndex = selectedFloor.getSpaces().indexOf(copiedSpace);
            if (newIndex >= 0) {
                radiationTab.selectSpaceByIndex(newIndex);
            }
        }
    }


    // Создаем глубокую копию помещения с новыми ID комнат
    private Space createSpaceCopyWithNewIds(Space originalSpace) {
        Space spaceCopy = new Space();
        spaceCopy.setIdentifier(originalSpace.getIdentifier());
        spaceCopy.setType(originalSpace.getType());

        for (Room originalRoom : originalSpace.getRooms()) {
            Room roomCopy = createRoomCopy(originalRoom);
// Освещение: при копировании всегда без галочки
            roomCopy.setSelected(false);
            spaceCopy.addRoom(roomCopy);
        }
        return spaceCopy;
    }

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

    // Операции с этажами
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

// <<< НОВОЕ: ставим позицию в КОНЕЦ секции >>>
            int maxPos = building.getFloors().stream()
                    .filter(f -> f.getSectionIndex() == secIdx)
                    .mapToInt(Floor::getPosition)
                    .max().orElse(-1);
            floor.setPosition(maxPos + 1);

            building.addFloor(floor);

            refreshFloorListForSelectedSection();
            updateRadiationTab(building, true);
            updateLightingTab(building, /*autoApplyDefaults=*/true);
            updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

        }
    }

    private void updateRadiationTab(Building building, boolean forceOfficeSelection) {
        // дефолт: старое поведение — авто-правила включены
        updateRadiationTab(building, forceOfficeSelection, true);
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

        // 1. Сохраняем ВСЕ текущие состояния комнат
        RadiationTab radiationTab = getRadiationTab();
        Map<Integer, Boolean> allRoomStates = new HashMap<>();
        if (radiationTab != null) {
            allRoomStates.putAll(radiationTab.globalRoomSelectionMap);
        }
        LightingTab lightingTab = getLightingTab();
        if (lightingTab != null) {
            lightingTab.updateRoomSelectionStates();
        }

        // 2. Создаем копию этажа
        Floor copiedFloor = createFloorCopy(selectedFloor);
        String newFloorNumber = generateNextFloorNumber(selectedFloor.getNumber(), selectedFloor.getSectionIndex());
        copiedFloor.setNumber(newFloorNumber);
        copiedFloor.setSectionIndex(selectedFloor.getSectionIndex());

// ВАЖНО: позицию ставим в КОНЕЦ секции → копия не первый этаж
        int secIdx = selectedFloor.getSectionIndex();
        int maxPos = building.getFloors().stream()
                .filter(f -> f.getSectionIndex() == secIdx)
                .mapToInt(Floor::getPosition)
                .max().orElse(-1);
        copiedFloor.setPosition(maxPos + 1);

        updateSpaceIdentifiers(copiedFloor, extractDigits(newFloorNumber));

        // 3. Добавляем новый этаж в модель
        building.addFloor(copiedFloor);
        floorListModel.addElement(copiedFloor);

        // 4. Обновляем ТОЛЬКО список этажей в UI
        floorList.setSelectedValue(copiedFloor, true);
// Освещение: при копировании этажа авто-правила НЕ применяем — галочки должны быть пустыми
        updateLightingTab(building, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        // 5. Восстанавливаем ВСЕ состояния комнат
        if (radiationTab != null) {
            radiationTab.globalRoomSelectionMap.clear();
            radiationTab.globalRoomSelectionMap.putAll(allRoomStates);

            // 6. Устанавливаем состояния для новых комнат
            for (int i = 0; i < selectedFloor.getSpaces().size(); i++) {
                Space origSpace = selectedFloor.getSpaces().get(i);
                Space copiedSpace = copiedFloor.getSpaces().get(i);

                for (int j = 0; j < origSpace.getRooms().size(); j++) {
                    Room origRoom = origSpace.getRooms().get(j);
                    Room copiedRoom = copiedSpace.getRooms().get(j);

                    Boolean state = allRoomStates.get(origRoom.getId());
                    if (state != null) {
                        radiationTab.setRoomSelectionState(copiedRoom.getId(), state);
                    }
                }
            }

            // 7. Обновляем UI RadiationTab
            radiationTab.refreshFloors();
        }
    }

    private String extractDigits(String input) {
        return input.replaceAll("\\D", ""); // Удаляем все не-цифры
    }

    // Новый: считает следующий номер ТОЛЬКО в пределах указанной секции.
    private String generateNextFloorNumber(String currentNumber, int sectionIndex) {
        Pattern p = Pattern.compile("\\d+");
        Matcher m = p.matcher(currentNumber);

        if (m.find()) {
            int baseNum = Integer.parseInt(m.group());
            String prefix = currentNumber.substring(0, m.start());
            String suffix = currentNumber.substring(m.end());

            // максимум среди этажей ТЕКУЩЕЙ секции
            int maxNumberInSection = building.getFloors().stream()
                    .filter(f -> f.getSectionIndex() == sectionIndex)
                    .map(Floor::getNumber)
                    .mapToInt(n -> {
                        Matcher mm = p.matcher(n);
                        return mm.find() ? Integer.parseInt(mm.group()) : Integer.MIN_VALUE;
                    })
                    .max()
                    .orElse(baseNum);

            int next = (maxNumberInSection != Integer.MIN_VALUE)
                    ? Math.max(baseNum, maxNumberInSection) + 1
                    : baseNum + 1;

            return prefix + next + suffix;
        }

        // Не нашли цифр — делаем уникальное имя в рамках секции
        return generateUniqueNonNumericNameWithinSection(currentNumber, sectionIndex);
    }

    // Сохранённый «обёртка»-вариант, если где-то ещё вызывается без секции — поведение как раньше.
    private String generateNextFloorNumber(String currentNumber) {
        return generateNextFloorNumber(currentNumber, -1);
    }

    // Уникальное имя для НЕчисловых этажей — только в пределах секции
    private String generateUniqueNonNumericNameWithinSection(String base, int sectionIndex) {
        Pattern pattern = Pattern.compile(Pattern.quote(base) + "(?: \\(копия (\\d+)\\))?");
        int maxCopy = building.getFloors().stream()
                .filter(f -> sectionIndex < 0 || f.getSectionIndex() == sectionIndex)
                .map(Floor::getNumber)
                .map(pattern::matcher)
                .filter(Matcher::matches)
                .mapToInt(m -> m.group(1) != null ? Integer.parseInt(m.group(1)) : 0)
                .max()
                .orElse(0);

        return base + (maxCopy == 0 ? "" : " (копия " + (maxCopy + 1) + ")");
    }

    private void updateSpaceIdentifiers(Floor floor, String newFloorDigits) {
        for (Space space : floor.getSpaces()) {
            String newId = updateIdentifier(space.getIdentifier(), newFloorDigits);
            space.setIdentifier(newId);
        }
    }

    private String updateIdentifier(String identifier, String newFloorDigits) {
        // Шаблон для поиска формата "X-Y" (например: "кв 1-1")
        Pattern pattern = Pattern.compile("(.*?)(\\d+)-(\\d+)(.*)");
        Matcher matcher = pattern.matcher(identifier);

        if (matcher.matches()) {
            String prefix = matcher.group(1);  // Префикс (например, "кв ")
            String roomNum = matcher.group(3);  // Номер помещения
            String suffix = matcher.group(4);   // Суффикс (если есть)

            // Формируем новый идентификатор: префикс + цифры этажа + номер помещения
            return prefix + newFloorDigits + "-" + roomNum + suffix;
        }
        return identifier; // Возвращаем оригинал, если паттерн не совпал
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
            // При смене типа на PUBLIC - создать помещение при необходимости
            if (floor.getType() == Floor.FloorType.PUBLIC) {
                createDefaultSpaceIfMissing(floor);
            }
            floorListModel.set(index, floor);
            updateSpaceList();
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);
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
        if (index >= 0) {
            building.getFloors().remove(index);
            floorListModel.remove(index);
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);
    }

    // Операции с помещениями
    private void addSpace(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) {
            showMessage("Выберите этаж!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        AddSpaceDialog dialog = new AddSpaceDialog((JFrame) SwingUtilities.getWindowAncestor(this), selectedFloor.getType());
        Space space = null; // Объявляем здесь
        if (dialog.showDialog()) {
            space = new Space(); // Инициализируем здесь
            space.setIdentifier(dialog.getSpaceIdentifier());
            space.setType(dialog.getSpaceType());
            selectedFloor.addSpace(space);
            spaceListModel.addElement(space);
        }

        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

        // Проверяем, что space был создан
        RadiationTab radiationTab = getRadiationTab();
        if (radiationTab != null && space != null) {
            int newIndex = selectedFloor.getSpaces().indexOf(space);
            if (newIndex >= 0) {
                radiationTab.selectSpaceByIndex(newIndex);
            }
        }
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
            spaceListModel.set(index, space);
            updateRoomList();
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

    }

    private void removeSpace(ActionEvent e) {
        Floor floor = floorList.getSelectedValue();
        int index = spaceList.getSelectedIndex();
        if (floor != null && index >= 0) {
            floor.getSpaces().remove(index);
            spaceListModel.remove(index);
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

    }


    // «умное» добавление комнат — с подсказками, пакетным добавлением и автонумерацией
    private void addRoom(ActionEvent e) {
        Space selectedSpace = spaceList.getSelectedValue();
        if (selectedSpace == null) {
            showMessage("Выберите помещение!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Только для квартир — быстрый набор через AddRoomDialog.
        if (isApartmentSpace(selectedSpace)) {
            java.util.List<String> suggestions = collectPopularApartmentRoomNames(building, 30);

            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            AddRoomDialog dlg = new AddRoomDialog(
                    parent,
                    suggestions,
                    "",     // пусто при добавлении
                    false
            );
            if (dlg.showDialog()) {
                String name = dlg.getNameToAdd();
                if (name != null && !name.isBlank()) {
                    Room room = new Room();
                    room.setName(name.trim());
                    selectedSpace.addRoom(room);
                    roomListModel.addElement(room);
                }
            }
        } else {
            // Для НЕ-квартир — простое поле ввода без «быстрого выбора»
            String name = showInputDialog("Добавление комнаты", "Название комнаты:", "");
            if (name != null && !name.isBlank()) {
                Room room = new Room();
                room.setName(name.trim());
                selectedSpace.addRoom(room);
                roomListModel.addElement(room);
            }
        }

        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);

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
            java.util.List<String> suggestions = collectPopularApartmentRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            AddRoomDialog dlg = new AddRoomDialog(
                    parent,
                    suggestions,
                    room.getName(),
                    true
            );
            if (dlg.showDialog()) {
                String newName = dlg.getNameToAdd();
                if (newName != null && !newName.isBlank()) {
                    room.setName(newName.trim());
                    roomListModel.set(index, room);
                    updateRadiationTab(building, true);
                    updateLightingTab(building, /*autoApplyDefaults=*/true);
                    updateMicroclimateTab(building, /*autoApplyDefaults=*/true);
                }
            }
        } else {
            // Для НЕ-квартир — простое редактирование строкой
            String newName = showInputDialog("Редактирование комнаты", "Новое название комнаты:", room.getName());
            if (newName != null && !newName.isBlank()) {
                room.setName(newName.trim());
                roomListModel.set(index, room);
                updateRadiationTab(building, true);
                updateLightingTab(building, /*autoApplyDefaults=*/true);
                updateMicroclimateTab(building, /*autoApplyDefaults=*/true);
            }
        }
    }


    private void removeRoom(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space != null && index >= 0) {
            space.getRooms().remove(index);
            roomListModel.remove(index);
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/true);
    }

    // Вспомогательные методы
    private void updateSpaceList() {
        spaceListModel.clear();
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) return;

        // просто показываем реальные помещения, ничего не добавляем автоматически
        List<Space> sorted = new ArrayList<>(selectedFloor.getSpaces());
        sorted.sort(Comparator.comparingInt(Space::getPosition));
        sorted.forEach(spaceListModel::addElement);

        if (!spaceListModel.isEmpty()) {
            // попытка сохранить текущее выделение (если есть)
            Space prev = spaceList.getSelectedValue();
            int idx = (prev != null) ? sorted.indexOf(prev) : -1;
            spaceList.setSelectedIndex(idx >= 0 ? idx : 0);
        }
    }

    private void createDefaultSpaceIfMissing(Floor floor) {
        boolean hasDefaultSpace = floor.getSpaces().stream()
                .anyMatch(space -> "-".equals(space.getIdentifier()));

        if (!hasDefaultSpace) {
            Space defaultSpace = new Space();
            defaultSpace.setIdentifier("-");
            defaultSpace.setType(Space.SpaceType.PUBLIC_SPACE); // Используем соответствующий тип
            floor.addSpace(defaultSpace);
        }
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

        // Помещения/комнаты для актуального выбора
        updateSpaceList();
        updateRoomList();
    }

    private void updateLightingTab(Building building, boolean autoApplyDefaults) {
        LightingTab tab = getLightingTab();
        if (tab != null) {
            tab.setBuilding(building, autoApplyDefaults);
            tab.refreshData();
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
        JPanel panel = new JPanel(new GridLayout(1, 2, 5, 5));
        panel.add(new JLabel(message));
        JTextField field = new JTextField(initialValue);
        panel.add(field);

        int result = JOptionPane.showConfirmDialog(
                this, panel, title, JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

        return (result == JOptionPane.OK_OPTION) ? field.getText() : null;
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
        updateRadiationTab(this.building, true);
        updateLightingTab(building, true);
        updateMicroclimateTab(building, true);

        if (floorList != null && !floorListModel.isEmpty()) {
            floorList.setSelectedIndex(0);
        }
        if (floorList != null) {
            floorList.addListSelectionListener(evt -> {
                if (!evt.getValueIsAdjusting()) updateSpaceList();
            });
        }
        if (spaceList != null) {
            spaceList.addListSelectionListener(evt -> {
                if (!evt.getValueIsAdjusting()) updateRoomList();
            });
        }
    }
    // Вспомогательная функция: топ часто встречающихся названий комнат в здании
    private java.util.List<String> collectPopularRoomNames(Building b, int limit) {
        java.util.Map<String, Integer> freq = new java.util.HashMap<>();
        if (b != null) {
            for (Floor f : b.getFloors()) {
                for (Space s : f.getSpaces()) {
                    for (Room r : s.getRooms()) {
                        String n = (r.getName() == null) ? "" : r.getName().trim();
                        if (!n.isEmpty()) freq.merge(n, 1, Integer::sum);
                    }
                }
            }
        }
        return freq.entrySet().stream()
                .sorted((a, c) -> {
                    int byCount = Integer.compare(c.getValue(), a.getValue());
                    if (byCount != 0) return byCount;
                    return a.getKey().compareToIgnoreCase(c.getKey());
                })
                .limit(limit)
                .map(java.util.Map.Entry::getKey)
                .toList();
    }

    private java.util.List<String> collectPopularApartmentRoomNames(Building b, int limit) {
        // key — нормализованная «база» (в нижнем регистре), val — счётчик
        java.util.Map<String, Integer> freq = new java.util.HashMap<>();
        // для отображения храним «красивый» вариант (первое встреченное написание базы)
        java.util.Map<String, String> display = new java.util.LinkedHashMap<>();

        if (b != null) {
            for (Floor f : b.getFloors()) {
                for (Space s : f.getSpaces()) {
                    if (!isApartmentSpace(s)) continue;
                    for (Room r : s.getRooms()) {
                        String raw = (r.getName() == null) ? "" : r.getName().trim();
                        if (raw.isEmpty()) continue;

                        String base = normalizeRoomBaseName(raw);
                        if (base.isEmpty()) continue;

                        String key = base.toLowerCase(java.util.Locale.ROOT);
                        freq.merge(key, 1, Integer::sum);
                        display.putIfAbsent(key, base); // запоминаем первое «красивое» написание
                    }
                }
            }
        }

        return freq.entrySet().stream()
                .sorted((a, c) -> {
                    int byCount = Integer.compare(c.getValue(), a.getValue());
                    if (byCount != 0) return byCount;
                    return display.get(a.getKey()).compareToIgnoreCase(display.get(c.getKey()));
                })
                .limit(limit)
                .map(e -> display.get(e.getKey()))
                .toList();
    }


    // Возвращает true, если помещение — «квартира».
// Не опираемся на конкретные enum-константы проекта, работаем по имени типа и по идентификатору.
    private boolean isApartmentSpace(Space s) {
        if (s == null) return false;

        // 1) По типу (через имя enum, чтобы не падать, если нет APARTMENT в исходном enum)
        Space.SpaceType t = s.getType();
        if (t != null) {
            String tn = t.name();
            if ("APARTMENT".equalsIgnoreCase(tn) ||
                    "FLAT".equalsIgnoreCase(tn) ||
                    "RESIDENTIAL".equalsIgnoreCase(tn) ||
                    "LIVING".equalsIgnoreCase(tn)) {
                return true;
            }
        }

        // 2) По идентификатору (часто «кв 1-1», «квартира 12» и т.п.)
        String id = s.getIdentifier();
        if (id != null) {
            String low = id.toLowerCase(java.util.Locale.ROOT);
            if (low.startsWith("кв") || low.contains("квартира")) return true;
            if (low.startsWith("apt") || low.contains("apartment")) return true;
        }
        return false;
    }
    // КОПИЯ ДЛЯ СОХРАНЕНИЯ: сохраняем id и selected
    private Floor createFloorCopyPreserve(Floor original) {
        Floor copy = new Floor();
        copy.setNumber(original.getNumber());
        copy.setType(original.getType());
        copy.setName(original.getName());
        copy.setSectionIndex(original.getSectionIndex());
        copy.setPosition(original.getPosition());
        for (Space os : original.getSpaces()) {
            copy.addSpace(createSpaceCopyPreserve(os));
        }
        return copy;
    }

    private Space createSpaceCopyPreserve(Space original) {
        Space copy = new Space();
        copy.setIdentifier(original.getIdentifier());
        copy.setType(original.getType());
        copy.setPosition(original.getPosition());
        // Сохраняем комнаты как есть (id + selected + наружные стены)
        for (Room or : original.getRooms()) {
            Room r = new Room();
            r.setId(or.getId());                   // сохраняем id
            r.setName(or.getName());
            r.setVolume(or.getVolume());
            r.setVentilationChannels(or.getVentilationChannels());
            r.setVentilationSectionArea(or.getVentilationSectionArea());
            r.setSelected(or.isSelected());        // сохраняем выбранность
            r.setOriginalRoomId(or.getOriginalRoomId());
            // НОВОЕ: переносим количество наружных стен, чтобы не терялось при сохранении
            try { r.setExternalWallsCount(or.getExternalWallsCount()); } catch (Throwable ignore) {}
            copy.addRoom(r);
        }
        return copy;
    }

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

}