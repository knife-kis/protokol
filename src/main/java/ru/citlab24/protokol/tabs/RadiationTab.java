package ru.citlab24.protokol.tabs;

import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;
import ru.citlab24.protokol.tabs.renderers.RoomListRenderer;
import ru.citlab24.protokol.db.DatabaseManager;

import javax.swing.*;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import java.awt.*;
import java.util.List;
import java.util.stream.Collectors;

public class RadiationTab extends JPanel {
    private JComboBox<Building> buildingComboBox;
    private JList<Floor> floorList;
    private JList<Room> roomList;
    private DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private DefaultListModel<Room> roomListModel = new DefaultListModel<>();
    private JTable dataTable;
    private RadiationTableModel tableModel;
    private JButton addRoomButton;
    private JButton removeRoomButton;
    private JButton exportButton;

    public RadiationTab() {
        setLayout(new BorderLayout());

        // Панель выбора здания и этажа
        JPanel selectionPanel = new JPanel(new GridLayout(1, 3));

        // Здания
        buildingComboBox = new JComboBox<>();
        buildingComboBox.addActionListener(e -> updateFloorList());
        selectionPanel.add(new JScrollPane(buildingComboBox));

        // Этажи
        floorList = new JList<>(floorListModel);
        floorList.setCellRenderer(new FloorListRenderer());
        floorList.addListSelectionListener(e -> updateRoomList());
        selectionPanel.add(new JScrollPane(floorList));

        // Помещения
        roomList = new JList<>(roomListModel);
        roomList.setCellRenderer(new RoomListRenderer());
        roomList.addListSelectionListener(e -> {
            if (e.getValueIsAdjusting()) return;

            Room selected = roomList.getSelectedValue();
            if (selected != null) {
                // Проверка на дублирование
                boolean alreadyExists = false;
                for (int i = 0; i < tableModel.getRowCount(); i++) {
                    RadiationRecord record = tableModel.getRecords().get(i);
                    if (record.getRoom().equals(selected.getName())) {
                        alreadyExists = true;
                        break;
                    }
                }

                if (!alreadyExists) {
                    RadiationRecord record = new RadiationRecord();
                    record.setRoom(selected.getName());

                    Floor selectedFloor = floorList.getSelectedValue();
                    if (selectedFloor != null) {
                        record.setFloor(selectedFloor.getName());
                    }

                    record.setPermissibleLevel(0.25);
                    tableModel.addRecord(record);
                } else {
                    JOptionPane.showMessageDialog(this,
                            "Это помещение уже добавлено в таблицу",
                            "Предупреждение",
                            JOptionPane.WARNING_MESSAGE);
                }
            }
        });
        selectionPanel.add(new JScrollPane(roomList));

        add(selectionPanel, BorderLayout.NORTH);

        // Таблица данных
        tableModel = new RadiationTableModel();
        dataTable = new JTable(tableModel);
        add(new JScrollPane(dataTable), BorderLayout.CENTER);

        // Панель кнопок
        JPanel buttonPanel = new JPanel();

        addRoomButton = new JButton("Добавить помещение");
        addRoomButton.addActionListener(e -> {
            // Можно добавить диалог для ввода нового помещения
            RadiationRecord record = new RadiationRecord();
            record.setPermissibleLevel(0.25); // Значение по умолчанию
            tableModel.addRecord(record);
        });

        removeRoomButton = new JButton("Удалить помещение");
        removeRoomButton.addActionListener(e -> {
            int selectedRow = dataTable.getSelectedRow();
            if (selectedRow != -1) {
                tableModel.removeRecord(selectedRow);
            }
        });

        exportButton = new JButton("Экспорт в Excel");
        exportButton.addActionListener(e -> RadiationExcelExporter.export(tableModel.getRecords(), this));

        buttonPanel.add(addRoomButton);
        buttonPanel.add(removeRoomButton);
        buttonPanel.add(exportButton);

        add(buttonPanel, BorderLayout.SOUTH);

        // Загрузка зданий
        loadBuildings();
    }

    private void loadBuildings() {
        // Загрузка зданий из базы данных
        List<Building> buildings = DatabaseManager.getBuildings();
        for (Building building : buildings) {
            buildingComboBox.addItem(building);
        }
    }

    private void updateFloorList() {
        floorListModel.clear();
        Building selectedBuilding = (Building) buildingComboBox.getSelectedItem();
        if (selectedBuilding != null) {
            List<Floor> floors = DatabaseManager.getFloors(selectedBuilding.getId());
            for (Floor floor : floors) {
                floorListModel.addElement(floor);
            }
        }
    }

    private void updateRoomList() {
        roomListModel.clear();
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor != null) {
            List<Room> rooms = DatabaseManager.getRooms(selectedFloor.getId());
            for (Room room : rooms) {
                roomListModel.addElement(room);
            }
        }
    }
}