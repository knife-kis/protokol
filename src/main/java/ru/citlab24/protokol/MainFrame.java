package ru.citlab24.protokol;

import ru.citlab24.protokol.tabs.buildingTab.BuildingTab;
import ru.citlab24.protokol.tabs.microclimateTab.MicroclimateTab;

import javax.swing.*;
import java.awt.*;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;

public class MainFrame extends JFrame {

    public MainFrame() {
        super("Building Analytics");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 750);
        setLocationRelativeTo(null);

        initUI();
    }

    private void initUI() {
        // Создаем панель с вкладками
        JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP, JTabbedPane.SCROLL_TAB_LAYOUT);
        tabbedPane.setFont(new Font("Segoe UI", Font.PLAIN, 14));

        // Добавляем вкладки с иконками
        tabbedPane.addTab("Характеристики здания",
                FontIcon.of(FontAwesomeSolid.BUILDING, 24, new Color(0, 115, 200)),
                new BuildingTab());

        tabbedPane.addTab("Микроклимат",
                FontIcon.of(FontAwesomeSolid.THERMOMETER_HALF, 24, new Color(76, 175, 80)),
                new MicroclimateTab());

        // УБРАНА ПАНЕЛЬ СТАТУСА (прогресс-бар и кнопка "Сохранить проект")

        // Основная компоновка
        setLayout(new BorderLayout());
        add(tabbedPane, BorderLayout.CENTER);
    }
}