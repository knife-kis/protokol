package ru.citlab24.protokol.tabs.buildingTab;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.sql.SQLException;
import java.util.List;
import java.util.Locale;

import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.VentilationTab;
import ru.citlab24.protokol.tabs.dialogs.AddFloorDialog;
import ru.citlab24.protokol.tabs.dialogs.AddSpaceDialog;
import ru.citlab24.protokol.tabs.dialogs.LoadProjectDialog;
import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.renderers.*;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;

public class BuildingTab extends JPanel {
    private Building building;
    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final DefaultListModel<Space> spaceListModel = new DefaultListModel<>();
    private final DefaultListModel<Room> roomListModel = new DefaultListModel<>();

    private JList<Floor> floorList;
    private JList<Space> spaceList;
    private JList<Room> roomList;

    public BuildingTab(Building building) {
        setRussianLocale();
        this.building = building;
        initComponents();
    }
    private void setRussianLocale() {
        Locale.setDefault(new Locale("ru", "RU"));
        UIManager.put("OptionPane.yesButtonText", "Да");
        UIManager.put("OptionPane.noButtonText", "Нет");
        UIManager.put("OptionPane.cancelButtonText", "Отмена");
        // Важно: обновляем ресурсы UIManager
        for (Object key : UIManager.getLookAndFeelDefaults().keySet()) {
            if (key instanceof String && ((String) key).endsWith(".locale")) {
                UIManager.getLookAndFeelDefaults().put(key, Locale.getDefault());
            }
        }
    }
    private void initComponents() {
        setLayout(new BorderLayout());
        add(createBuildingPanel(), BorderLayout.CENTER);
        add(createActionButtons(), BorderLayout.SOUTH);
    }
    private JPanel createBuildingPanel() {
        JPanel panel = new JPanel(new GridLayout(1, 3, 15, 15));
        panel.add(createFloorPanel());
        panel.add(createSpacePanel());
        panel.add(createRoomPanel());
        return panel;
    }
    private JPanel createFloorPanel() {
        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(BorderFactory.createTitledBorder(null,
                "Этажи здания", TitledBorder.LEFT, TitledBorder.TOP,
                new Font("Segoe UI", Font.BOLD, 14), new Color(0,115,200)));

        floorList = new JList<>(floorListModel);
        floorList.setCellRenderer(new FloorListRenderer());
        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        floorList.addListSelectionListener(e -> updateSpaceList());

        JPanel btns = new JPanel(new GridLayout(1, 3, 5, 5));
        btns.add(createStyledButton("Добавить", FontAwesomeSolid.PLUS_CIRCLE, new Color(46,125,50), this::addFloor));
        btns.add(createStyledButton("Изменить", FontAwesomeSolid.EDIT, new Color(255, 152, 0), this::editFloor));
        btns.add(createStyledButton("Удалить", FontAwesomeSolid.TRASH, new Color(198,40,40), this::removeFloor));

        p.add(new JScrollPane(floorList), BorderLayout.CENTER);
        p.add(btns, BorderLayout.SOUTH);
        return p;
    }

    // ДОБАВЛЕННЫЙ МЕТОД: Панель с кнопками действий
    private JPanel createActionButtons() {
        JPanel p = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        p.add(createStyledButton("Загрузить проект", FontAwesomeSolid.FOLDER_OPEN, new Color(33, 150, 243), e -> loadProject()));
        p.add(createStyledButton("Сохранить проект", FontAwesomeSolid.SAVE, new Color(0,115,200), e -> saveProject()));
        p.add(createStyledButton("Рассчитать показатели", FontAwesomeSolid.CALCULATOR, new Color(103,58,183), e -> calculateMetrics()));
        return p;
    }

    private void loadProject() {
        try {
            List<Building> projects = DatabaseManager.getAllBuildings();

            if (projects.isEmpty()) {
                JOptionPane.showMessageDialog(
                        this,
                        "Нет сохраненных проектов",
                        "Информация",
                        JOptionPane.INFORMATION_MESSAGE
                );
                return;
            }

            LoadProjectDialog dialog = new LoadProjectDialog(
                    (JFrame) SwingUtilities.getWindowAncestor(this),
                    projects
            );
            dialog.setVisible(true);

            Building selectedProject = dialog.getSelectedProject();
            if (selectedProject != null) {
                Building loadedBuilding = DatabaseManager.loadBuilding(selectedProject.getId());
                this.building = loadedBuilding;
                refreshAllLists();
                JOptionPane.showMessageDialog(
                        this,
                        "Проект '" + loadedBuilding.getName() + "' успешно загружен",
                        "Загрузка",
                        JOptionPane.INFORMATION_MESSAGE
                );
            }
        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(
                    this,
                    "Ошибка загрузки проектов: " + e.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }

    private void refreshAllLists() {
        floorListModel.clear();
        spaceListModel.clear();
        roomListModel.clear();

        for (Floor floor : building.getFloors()) {
            floorListModel.addElement(floor);
        }

        // Обновляем выбранные элементы
        if (!building.getFloors().isEmpty()) {
            floorList.setSelectedIndex(0);
        }
    }

    private void saveProject() {
        try {
            String name = JOptionPane.showInputDialog(
                    this,
                    "Введите название проекта:",
                    "Сохранение проекта",
                    JOptionPane.PLAIN_MESSAGE
            );

            if (name != null && !name.trim().isEmpty()) {
                building.setName(name.trim());
                DatabaseManager.saveBuilding(building);
                JOptionPane.showMessageDialog(
                        this,
                        "Проект успешно сохранен",
                        "Сохранение",
                        JOptionPane.INFORMATION_MESSAGE
                );
            }
        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(
                    this,
                    "Ошибка сохранения проекта: " + e.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }

    private void calculateMetrics() {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            // Получаем доступ к вкладке вентиляции через главное окно
            Component[] tabs = ((MainFrame) mainFrame).getTabbedPane().getComponents();
            for (Component tab : tabs) {
                if (tab instanceof VentilationTab) {
                    ((VentilationTab) tab).refreshData();
                    break;
                }
            }

            // Переключаемся на вкладку вентиляции
            ((MainFrame) mainFrame).selectVentilationTab();
        }

        JOptionPane.showMessageDialog(this, "Данные для вентиляции обновлены!",
                "Расчет завершен", JOptionPane.INFORMATION_MESSAGE);
    }
    private JButton createActionButton(String text, Color bgColor, ActionListener action) {
        JButton btn = new JButton(text);
        btn.setBackground(bgColor);
        btn.setForeground(Color.WHITE);
        btn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        btn.addActionListener(action);
        return btn;
    }
    private JPanel createSpacePanel() {
        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(BorderFactory.createTitledBorder(null,
                "Помещения на этаже", TitledBorder.LEFT, TitledBorder.TOP,
                new Font("Segoe UI", Font.BOLD, 14), new Color(76,175,80)));

        spaceList = new JList<>(spaceListModel);
        spaceList.setCellRenderer(new SpaceListRenderer());
        spaceList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceList.addListSelectionListener(e -> updateRoomList());

        JPanel btns = new JPanel(new GridLayout(1, 3, 5, 5)); // Изменено на 3 кнопки
        btns.add(createStyledButton("Добавить", FontAwesomeSolid.PLUS, new Color(46,125,50), this::addSpace));
        btns.add(createStyledButton("Изменить", FontAwesomeSolid.EDIT, new Color(255, 152, 0), this::editSpace));
        btns.add(createStyledButton("Удалить", FontAwesomeSolid.TRASH, new Color(198,40,40), this::removeSpace));
        p.add(new JScrollPane(spaceList), BorderLayout.CENTER);
        p.add(btns, BorderLayout.SOUTH);
        return p;
    }

    private JPanel createRoomPanel() {
        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(BorderFactory.createTitledBorder(null,
                "Комнаты в помещении", TitledBorder.LEFT, TitledBorder.TOP,
                new Font("Segoe UI", Font.BOLD, 14), new Color(156,39,176)));

        roomList = new JList<>(roomListModel);
        roomList.setCellRenderer(new RoomListRenderer());

        JPanel btns = new JPanel(new GridLayout(1, 3, 5, 5));
        btns.add(createStyledButton("Добавить", FontAwesomeSolid.PLUS_SQUARE, new Color(46,125,50), this::addRoom));
        btns.add(createStyledButton("Изменить", FontAwesomeSolid.EDIT, new Color(255, 152, 0), this::editRoom));
        btns.add(createStyledButton("Удалить", FontAwesomeSolid.TRASH, new Color(198,40,40), this::removeRoom));

        p.add(new JScrollPane(roomList), BorderLayout.CENTER);
        p.add(btns, BorderLayout.SOUTH);
        return p;
    }

    private JPanel createButtonPanel() {
        JPanel p = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        p.add(createStyledButton("Сохранить проект", FontAwesomeSolid.SAVE, new Color(0,115,200), e -> {/*...*/}));
        p.add(createStyledButton("Рассчитать показатели", FontAwesomeSolid.CALCULATOR, new Color(103,58,183), e -> {/*...*/}));
        return p;
    }

    private JButton createStyledButton(String text,
                                       FontAwesomeSolid icon,
                                       Color bgColor,
                                       java.awt.event.ActionListener action) {
        JButton btn = new JButton(text, FontIcon.of(icon, 16, Color.WHITE));
        btn.setBackground(bgColor);
        btn.setForeground(Color.WHITE);
        btn.setFocusPainted(false);
        btn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        btn.addActionListener(action);
        btn.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) { btn.setBackground(bgColor.darker()); }
            public void mouseExited(java.awt.event.MouseEvent e) { btn.setBackground(bgColor); }
        });
        return btn;
    }
    // Обработчики событий
    private void addFloor(ActionEvent e) {
        AddFloorDialog dialog = new AddFloorDialog((JFrame) SwingUtilities.getWindowAncestor(this));
        if (dialog.showDialog()) {
            Floor floor = new Floor();
            // ИСПРАВЛЕНО: используем setNumber вместо setName
            floor.setNumber(dialog.getFloorNumber());
            floor.setType(dialog.getFloorType());
            building.addFloor(floor);
            floorListModel.addElement(floor);
        }
    }

    private void removeFloor(ActionEvent e) {
        int index = floorList.getSelectedIndex();
        if (index >= 0) {
            building.getFloors().remove(index);
            floorListModel.remove(index);
        }
    }

    private void addSpace(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) {
            JOptionPane.showMessageDialog(this, "Выберите этаж!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        AddSpaceDialog dialog = new AddSpaceDialog((JFrame) SwingUtilities.getWindowAncestor(this), selectedFloor.getType());
        if (dialog.showDialog()) {
            Space space = new Space();
            space.setIdentifier(dialog.getSpaceIdentifier());
            space.setType(dialog.getSpaceType());
            selectedFloor.addSpace(space);
            spaceListModel.addElement(space);
        }
    }

    private void addRoom(ActionEvent e) {
        Space selectedSpace = spaceList.getSelectedValue();
        if (selectedSpace == null) {
            JOptionPane.showMessageDialog(this, "Выберите помещение!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Упрощенная панель с одним полем (только название)
        JPanel panel = new JPanel(new GridLayout(1, 2, 5, 5));
        panel.add(new JLabel("Название комнаты:"));
        JTextField nameField = new JTextField();
        panel.add(nameField);

        int result = JOptionPane.showConfirmDialog(this, panel, "Добавление комнаты",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

        if (result == JOptionPane.OK_OPTION) {
            String name = nameField.getText().trim();

            if (!name.isEmpty()) {
                Room room = new Room();
                room.setName(name);
                selectedSpace.addRoom(room);
                roomListModel.addElement(room);
            } else {
                JOptionPane.showMessageDialog(this, "Введите название комнаты!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            }
        }
    }

    private void updateSpaceList() {
        spaceListModel.clear();
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor != null) {
            for (Space space : selectedFloor.getSpaces()) {
                spaceListModel.addElement(space);
            }
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
    private void removeSpace(ActionEvent e) {
        Floor floor = floorList.getSelectedValue();
        int index = spaceList.getSelectedIndex();

        if (floor != null && index >= 0) {
            floor.getSpaces().remove(index);
            spaceListModel.remove(index);
        }
    }

    private void removeRoom(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space != null && index >= 0) {
            space.getRooms().remove(index);
            roomListModel.remove(index);
        }
    }
    private void editFloor(ActionEvent e) {
        int index = floorList.getSelectedIndex();
        if (index < 0) {
            JOptionPane.showMessageDialog(this, "Выберите этаж для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
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
    }
    private void editSpace(ActionEvent e) {
        Floor floor = floorList.getSelectedValue();
        int index = spaceList.getSelectedIndex();

        if (floor == null || index < 0) {
            JOptionPane.showMessageDialog(this, "Выберите помещение для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Space space = spaceListModel.get(index);
        AddSpaceDialog dialog = new AddSpaceDialog((JFrame) SwingUtilities.getWindowAncestor(this), floor.getType());
        dialog.setSpaceIdentifier(space.getIdentifier());
        dialog.setSpaceType(space.getType());

        if (dialog.showDialog()) {
            space.setIdentifier(dialog.getSpaceIdentifier());
            space.setType(dialog.getSpaceType());
            spaceListModel.set(index, space); // Обновляем элемент
            updateRoomList(); // Обновляем зависимые списки
        }
    }

    private void editRoom(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space == null || index < 0) {
            JOptionPane.showMessageDialog(this, "Выберите комнату для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Room room = roomListModel.get(index);
        String newName = JOptionPane.showInputDialog(this, "Введите новое название комнаты:", room.getName());

        if (newName != null && !newName.trim().isEmpty()) {
            room.setName(newName.trim());
            roomListModel.set(index, room); // Обновляем элемент
        }
    }
}