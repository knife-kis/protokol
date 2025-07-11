package ru.citlab24.protokol.tabs.buildingTab;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.util.List;

import ru.citlab24.protokol.tabs.dialogs.AddFloorDialog;
import ru.citlab24.protokol.tabs.dialogs.AddSpaceDialog;
import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.renderers.*;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;

public class BuildingTab extends JPanel {
    private final Building building = new Building();
    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final DefaultListModel<Space> spaceListModel = new DefaultListModel<>();
    private final DefaultListModel<Room> roomListModel = new DefaultListModel<>();

    private JList<Floor> floorList;
    private JList<Space> spaceList;
    private JList<Room> roomList;

    public BuildingTab() {
        setLayout(new BorderLayout(10, 10));
        setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));
        initComponents();
    }

    private void initComponents() {
        JSplitPane mainSplit = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
                createFloorPanel(),
                new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
                        createSpacePanel(),
                        createRoomPanel()));
        mainSplit.setDividerLocation(0.3);
        add(mainSplit, BorderLayout.CENTER);
        add(createButtonPanel(), BorderLayout.SOUTH);
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

        JPanel btns = new JPanel(new GridLayout(1,2,5,5));
        btns.add(createStyledButton("Добавить этаж", FontAwesomeSolid.PLUS_CIRCLE, new Color(46,125,50), this::addFloor));
        btns.add(createStyledButton("Удалить этаж", FontAwesomeSolid.TRASH, new Color(198,40,40), this::removeFloor));


        p.add(new JScrollPane(floorList), BorderLayout.CENTER);
        p.add(btns, BorderLayout.SOUTH);
        return p;
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

        JPanel btns = new JPanel(new GridLayout(1, 2, 5, 5));
        btns.add(createStyledButton("Добавить помещение", FontAwesomeSolid.PLUS, new Color(46,125,50), this::addSpace));
        btns.add(createStyledButton("Удалить помещение", FontAwesomeSolid.TRASH, new Color(198,40,40), this::removeSpace));
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

        JPanel btns = new JPanel(new GridLayout(1, 2, 5, 5));
        btns.add(createStyledButton("Добавить комнату", FontAwesomeSolid.PLUS_SQUARE, new Color(46,125,50), this::addRoom));
        btns.add(createStyledButton("Удалить комнату", FontAwesomeSolid.TRASH, new Color(198,40,40), this::removeRoom)); // Добавлено

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

        String roomName = JOptionPane.showInputDialog(this, "Введите название комнаты:", "Добавление комнаты", JOptionPane.PLAIN_MESSAGE);
        if (roomName != null && !roomName.trim().isEmpty()) {
            Room room = new Room();
            room.setName(roomName.trim());
            selectedSpace.addRoom(room);
            roomListModel.addElement(room);
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
}