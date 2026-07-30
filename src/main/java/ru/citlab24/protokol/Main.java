package ru.citlab24.protokol;

import ru.citlab24.protokol.db.DatabaseConfig;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.db.AppUserRecord;

import javax.swing.*;
import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

public class Main {
    public static void main(String[] args) {
        Locale.setDefault(new Locale("ru", "RU"));
        SwingUtilities.invokeLater(() -> {
            SwingUtilities.invokeLater(() -> {
                AppTheme.install();
                try {
                    DatabaseManager.initialize();
                } catch (Throwable error) {
                    showStartupError("Подключение к базе данных", error,
                            "Не удалось подключиться к общей базе PostgreSQL.\n" +
                                    "Проверьте службу PostgreSQL и файл:\n" + DatabaseConfig.getConfigPath() +
                                    "\n\nПричина: " + rootCause(error).getMessage());
                    return;
                }
                AppUserRecord currentUser;
                try {
                    currentUser = DatabaseManager.registerCurrentWindowsUser();
                } catch (Exception error) {
                    showStartupError("Вход пользователя", error,
                            "Не удалось зарегистрировать текущего пользователя Windows.\n\n" +
                                    "Причина: " + rootCause(error).getMessage());
                    return;
                }
                try {
                    javafx.application.Platform.setImplicitExit(false);
                    MainFrame frame = new MainFrame(currentUser);
                    frame.setLocationByPlatform(true);
                    frame.setVisible(true);
                } catch (Throwable error) {
                    showStartupError("Открытие главного окна", error,
                            "Не удалось открыть главное окно программы.\n\nПричина: "
                                    + rootCause(error).getMessage());
                }
            });
        });
    }

    private static void showStartupError(String stage, Throwable error, String message) {
        Path logFile = writeStartupLog(stage, error);
        String logText = logFile == null
                ? "Журнал запуска записать не удалось."
                : "Подробный журнал:\n" + logFile;
        String diagnostic = message + "\n\n" + logText;
        Object[] options = {"Открыть папку журнала", "Скопировать причину", "Закрыть"};
        int result = JOptionPane.showOptionDialog(
                null, diagnostic, "Ошибка запуска программы",
                JOptionPane.DEFAULT_OPTION, JOptionPane.ERROR_MESSAGE,
                null, options, options[2]);
        if (result == 0 && logFile != null) {
            try {
                Desktop.getDesktop().open(logFile.getParent().toFile());
            } catch (Exception ignored) {
            }
        } else if (result == 1) {
            Toolkit.getDefaultToolkit().getSystemClipboard()
                    .setContents(new StringSelection(diagnostic), null);
        }
    }

    private static Path writeStartupLog(String stage, Throwable error) {
        StringWriter stackTrace = new StringWriter();
        error.printStackTrace(new PrintWriter(stackTrace));
        String entry = "\n============================================================\n"
                + LocalDateTime.now().format(
                DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm:ss"))
                + " · " + stage + "\n"
                + stackTrace;
        Path primary = DatabaseConfig.getConfigPath().getParent()
                .resolve("logs").resolve("startup.log");
        Path fallback = Path.of(System.getProperty("java.io.tmpdir"),
                "CITLAB24-Protokol-startup.log");
        for (Path path : new Path[]{primary, fallback}) {
            try {
                Files.createDirectories(path.getParent());
                Files.writeString(path, entry, StandardCharsets.UTF_8,
                        StandardOpenOption.CREATE, StandardOpenOption.APPEND);
                return path;
            } catch (Exception ignored) {
            }
        }
        return null;
    }

    private static Throwable rootCause(Throwable error) {
        Throwable result = error;
        while (result.getCause() != null && result.getCause() != result) {
            result = result.getCause();
        }
        return result;
    }
}
