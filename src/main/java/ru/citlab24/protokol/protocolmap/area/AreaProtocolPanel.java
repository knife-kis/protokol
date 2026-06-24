package ru.citlab24.protokol.protocolmap.area;

import com.github.lgooddatepicker.components.DatePicker;
import com.github.lgooddatepicker.components.DatePickerSettings;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.titleTab.TechnicalAssignmentImporter;
import ru.citlab24.protokol.tabs.titleTab.TitlePageImportData;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;
import java.awt.image.BufferedImage;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.List;
import java.util.Locale;
import java.util.prefs.Preferences;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class AreaProtocolPanel extends JPanel {
    private static final String DATE_PATTERN = "dd.MM.yyyy";
    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern(DATE_PATTERN);
    private static final String CUSTOMER_PLACEHOLDER = "Название, почта, телефон";
    private static final String PREF_SKETCH_DIR = "areaProtocolSketchDir";
    private static final Preferences PREFS = Preferences.userNodeForPackage(AreaProtocolPanel.class);
    private static final String NOISE_METHOD_MI_LABEL = "МИ Ш.13-2021";
    private static final String NOISE_METHOD_ECOFIZIKA_LABEL = "РЭ Экофизика-110А";
    private static final String NOISE_METHOD_MI = "МИ Ш.13-2021 \"Методика измерений шума, инфразвука, воздушного\n"
            + "ультразвука на рабочих местах, в том числе рабочих местах\n"
            + "транспорта и объектов транспортной инфраструктуры, в\n"
            + "помещениях жилых, общественных и производственных\n"
            + "зданий, на селитебной и открытой территории.\" п. 12.3.3 / измерение шума";
    private static final String NOISE_METHOD_ECOFIZIKA = "Руководство по эксплуатации ПКДУ.411000.001.02 РЭ Шумомер-виброметр, "
            + "анализатор спектра Экофизика-110А п.7.1, п.7.2 (МИ ПКФ-12-006 методика измерений приложение "
            + "к руководству по эксплуатации п.2) / измерение шума";

    private final JTextField projectNameField = new JTextField(40);
    private final DatePicker protocolDatePicker;
    private final JTextField customerNameContactsField;
    private final JTextField customerLegalAddressField;
    private final JTextField customerActualAddressField;
    private final JTextArea objectNameArea;
    private final JTextField objectAddressField;
    private final JTextField contractNumberField;
    private final DatePicker contractDatePicker;
    private final JTextField applicationNumberField;
    private final DatePicker applicationDatePicker;
    private final JTextField representativeField;
    private final JTextField areaField = new JTextField();
    private final JTextField gammaMinValueField = new JTextField(8);
    private final JTextField gammaMaxValueField = new JTextField(8);
    private final JTextField gammaAverageValueField = new JTextField(8);
    private final JTextField pprMinValueField = new JTextField(8);
    private final JTextField pprMaxValueField = new JTextField(8);
    private final JTextField noiseEquivalentMinValueField = new JTextField(8);
    private final JTextField noiseEquivalentMaxValueField = new JTextField(8);
    private final JTextField noiseMaxLevelMinValueField = new JTextField(8);
    private final JTextField noiseMaxLevelMaxValueField = new JTextField(8);
    private final JComboBox<String> noiseMethodComboBox = new JComboBox<>(new String[]{NOISE_METHOD_MI_LABEL, NOISE_METHOD_ECOFIZIKA_LABEL});
    private final JPanel medValuesRowsPanel = new JPanel(new GridBagLayout());
    private final JCheckBox applyMedCountToAllCheckBox = new JCheckBox("Применить расстояние/количество ко всем профилям");
    private final List<MedValueRow> medValueRows = new ArrayList<>();
    private final JCheckBox medCheckBox = new JCheckBox("МЭД", true);
    private final JCheckBox pprCheckBox = new JCheckBox("ППР", true);
    private final JLabel imageLabel = new JLabel("Картинка не выбрана");
    private final SketchPreviewPanel sketchPreviewPanel = new SketchPreviewPanel();
    private JSlider gammaControlPointSizeSlider;
    private JSlider pprPointSizeSlider;
    private JSlider noisePointSizeSlider;
    private final JPanel measurementRowsPanel;
    private final List<MeasurementRow> measurementRows = new ArrayList<>();
    private final Color defaultTextColor = UIManager.getColor("TextField.foreground");
    private File imageFile;
    private boolean updatingMedValues;

    private static class MeasurementRow {
        JPanel panel;
        DatePicker datePicker;
        JTextField tempInsideStart;
        JTextField tempInsideEnd;
        JTextField tempOutsideStart;
        JTextField tempOutsideEnd;
        JCheckBox gammaCheckBox;
        JCheckBox pprCheckBox;
        JCheckBox noiseCheckBox;
    }

    private static class MedValueRow {
        int profileNumber;
        JTextField distanceField;
        JTextField countField;
    }

    private static final class AreaProjectSnapshot implements Serializable {
        private static final long serialVersionUID = 1L;

        String protocolDate;
        String projectName;
        String customerNameAndContacts;
        String customerLegalAddress;
        String customerActualAddress;
        String objectName;
        String objectAddress;
        String contractNumber;
        String contractDate;
        String applicationNumber;
        String applicationDate;
        String representative;
        String areaText;
        String gammaMinValue;
        String gammaMaxValue;
        String gammaAverageValue;
        String pprMinValue;
        String pprMaxValue;
        String noiseEquivalentMinValue;
        String noiseEquivalentMaxValue;
        String noiseMaxLevelMinValue;
        String noiseMaxLevelMaxValue;
        String noiseMethod;
        boolean medSelected;
        boolean pprSelected;
        String imageName;
        byte[] imageBytes;
        int boundaryColorRgb = Color.BLACK.getRGB();
        int gammaControlPointScale = 100;
        int pprPointScale = 100;
        int noisePointScale = 100;
        List<MedValueSnapshot> medValues = new ArrayList<>();
        List<MeasurementProjectRow> measurementRows = new ArrayList<>();
        List<PointSnapshot> boundaryPoints = new ArrayList<>();
        List<GammaProfileSnapshot> gammaProfiles = new ArrayList<>();
        List<GammaControlPointSnapshot> gammaControlPoints = new ArrayList<>();
        List<PprPointSnapshot> pprPoints = new ArrayList<>();
        List<NoisePointSnapshot> noisePoints = new ArrayList<>();
    }

    private static final class MedValueSnapshot implements Serializable {
        private static final long serialVersionUID = 1L;

        int profileNumber;
        String minValue;
        String maxValue;
        String distance;
        String count;
    }

    private static final class MeasurementProjectRow implements Serializable {
        private static final long serialVersionUID = 1L;

        String date;
        String tempInsideStart;
        String tempInsideEnd;
        String tempOutsideStart;
        String tempOutsideEnd;
        boolean gammaSelected;
        boolean pprSelected;
        boolean noiseSelected;
    }

    private static final class PointSnapshot implements Serializable {
        private static final long serialVersionUID = 1L;

        double x;
        double y;

        PointSnapshot(double x, double y) {
            this.x = x;
            this.y = y;
        }
    }

    private static final class GammaProfileSnapshot implements Serializable {
        private static final long serialVersionUID = 1L;

        double x;
        double y;
        double angle;
        String number;

        GammaProfileSnapshot(double x, double y, double angle, String number) {
            this.x = x;
            this.y = y;
            this.angle = angle;
            this.number = number;
        }
    }

    private static final class GammaControlPointSnapshot implements Serializable {
        private static final long serialVersionUID = 1L;

        double x;
        double y;
        String number;

        GammaControlPointSnapshot(double x, double y, String number) {
            this.x = x;
            this.y = y;
            this.number = number;
        }
    }

    private static final class PprPointSnapshot implements Serializable {
        private static final long serialVersionUID = 1L;

        double x;
        double y;
        String number;

        PprPointSnapshot(double x, double y, String number) {
            this.x = x;
            this.y = y;
            this.number = number;
        }
    }

    private static final class NoisePointSnapshot implements Serializable {
        private static final long serialVersionUID = 1L;

        double x;
        double y;
        String number;

        NoisePointSnapshot(double x, double y, String number) {
            this.x = x;
            this.y = y;
            this.number = number;
        }
    }

    private static class SketchPreviewPanel extends JPanel implements Scrollable {
        private BufferedImage image;
        private final List<Point2D.Double> boundaryPoints = new ArrayList<>();
        private final List<GammaProfile> gammaProfiles = new ArrayList<>();
        private final List<GammaControlPoint> gammaControlPoints = new ArrayList<>();
        private final List<PprPoint> pprPoints = new ArrayList<>();
        private final List<NoisePoint> noisePoints = new ArrayList<>();
        private final Deque<Runnable> undoActions = new ArrayDeque<>();
        private final Rectangle imageBounds = new Rectangle();
        private boolean boundaryMode;
        private boolean gammaProfileMode;
        private boolean gammaControlPointMode;
        private boolean pprPointMode;
        private boolean noisePointMode;
        private boolean radiationLayerVisible = true;
        private boolean noiseLayerVisible = true;
        private int gammaControlPointScale = 100;
        private int pprPointScale = 100;
        private int noisePointScale = 100;
        private Color boundaryColor = Color.BLACK;
        private Point currentMousePoint;
        private GammaProfile activeGammaProfile;
        private GammaProfile selectedGammaProfile;
        private GammaControlPoint activeGammaControlPoint;
        private GammaControlPoint selectedGammaControlPoint;
        private GammaControlPoint hoveredGammaControlPoint;
        private PprPoint activePprPoint;
        private PprPoint selectedPprPoint;
        private PprPoint hoveredPprPoint;
        private NoisePoint activeNoisePoint;
        private NoisePoint selectedNoisePoint;
        private NoisePoint hoveredNoisePoint;
        private Runnable boundaryModeOffAction;
        private Runnable gammaProfilesChangedAction = () -> {
        };
        private static final int GAMMA_RADIUS = 16;
        private static final int GAMMA_LINE_LENGTH = 44;
        private static final double GAMMA_ANGLE_STEP = Math.toRadians(7.5d);
        private static final int BOUNDARY_SNAP_RADIUS = 14;
        private static final int PPR_GAMMA_MIN_DISTANCE = 30;

        private static class GammaProfile {
            double x;
            double y;
            double angle;
            String number;

            GammaProfile(double x, double y, double angle) {
                this.x = x;
                this.y = y;
                this.angle = angle;
                this.number = "";
            }
        }

        private static class GammaControlPoint {
            double x;
            double y;
            String number;

            GammaControlPoint(double x, double y, String number) {
                this.x = x;
                this.y = y;
                this.number = number;
            }
        }

        private static class PprPoint {
            double x;
            double y;
            String number;

            PprPoint(double x, double y, String number) {
                this.x = x;
                this.y = y;
                this.number = number;
            }
        }

        private static class NoisePoint {
            double x;
            double y;
            String number;

            NoisePoint(double x, double y, String number) {
                this.x = x;
                this.y = y;
                this.number = number;
            }
        }

        SketchPreviewPanel() {
            setBackground(Color.WHITE);
            setPreferredSize(new Dimension(900, 560));
            setFocusable(true);

            MouseAdapter mouseHandler = new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent e) {
                    if (boundaryMode && SwingUtilities.isRightMouseButton(e)) {
                        if (boundaryModeOffAction != null) {
                            boundaryModeOffAction.run();
                        } else {
                            setBoundaryMode(false);
                        }
                        return;
                    }
                    if (image == null || !SwingUtilities.isLeftMouseButton(e)) {
                        return;
                    }
                    requestFocusInWindow();
                    if (noisePointMode) {
                        NoisePoint point = findNoisePoint(e.getPoint());
                        if (point != null) {
                            activeNoisePoint = point;
                            selectedNoisePoint = point;
                            selectedPprPoint = null;
                            selectedGammaControlPoint = null;
                            selectedGammaProfile = null;
                            repaint();
                            return;
                        }
                        Point2D.Double imagePoint = toImagePoint(e.getPoint());
                        if (imagePoint == null || !isInsideBoundary(imagePoint)) {
                            Toolkit.getDefaultToolkit().beep();
                            return;
                        }
                        selectedNoisePoint = new NoisePoint(imagePoint.x, imagePoint.y, "");
                        selectedPprPoint = null;
                        selectedGammaControlPoint = null;
                        selectedGammaProfile = null;
                        noisePoints.add(selectedNoisePoint);
                        NoisePoint addedPoint = selectedNoisePoint;
                        undoActions.push(() -> {
                            noisePoints.remove(addedPoint);
                            if (selectedNoisePoint == addedPoint) {
                                selectedNoisePoint = null;
                            }
                            if (hoveredNoisePoint == addedPoint) {
                                hoveredNoisePoint = null;
                            }
                            renumberNoisePoints();
                            repaint();
                        });
                        renumberNoisePoints();
                        repaint();
                        return;
                    }
                    if (pprPointMode) {
                        PprPoint point = findPprPoint(e.getPoint());
                        if (point != null) {
                            activePprPoint = point;
                            selectedPprPoint = point;
                            selectedNoisePoint = null;
                            selectedGammaControlPoint = null;
                            selectedGammaProfile = null;
                            repaint();
                            return;
                        }
                        Point2D.Double imagePoint = toImagePoint(e.getPoint());
                        if (imagePoint == null || !isInsideBoundary(imagePoint) || !isFarFromGammaMarks(imagePoint)) {
                            Toolkit.getDefaultToolkit().beep();
                            return;
                        }
                        selectedPprPoint = new PprPoint(imagePoint.x, imagePoint.y, "");
                        selectedNoisePoint = null;
                        selectedGammaControlPoint = null;
                        selectedGammaProfile = null;
                        pprPoints.add(selectedPprPoint);
                        PprPoint addedPoint = selectedPprPoint;
                        undoActions.push(() -> {
                            pprPoints.remove(addedPoint);
                            if (selectedPprPoint == addedPoint) {
                                selectedPprPoint = null;
                            }
                            if (hoveredPprPoint == addedPoint) {
                                hoveredPprPoint = null;
                            }
                            renumberPprPoints();
                            repaint();
                        });
                        renumberPprPoints();
                        repaint();
                        return;
                    }
                    if (gammaControlPointMode) {
                        GammaControlPoint point = findGammaControlPoint(e.getPoint());
                        if (point != null) {
                            activeGammaControlPoint = point;
                            selectedGammaControlPoint = point;
                            selectedNoisePoint = null;
                            selectedGammaProfile = null;
                            selectedPprPoint = null;
                            repaint();
                            return;
                        }
                        Point2D.Double imagePoint = toImagePoint(e.getPoint());
                        if (imagePoint == null || !isInsideBoundary(imagePoint)) {
                            return;
                        }
                        selectedGammaControlPoint = new GammaControlPoint(imagePoint.x, imagePoint.y, "");
                        selectedNoisePoint = null;
                        selectedGammaProfile = null;
                        selectedPprPoint = null;
                        gammaControlPoints.add(selectedGammaControlPoint);
                        GammaControlPoint addedPoint = selectedGammaControlPoint;
                        undoActions.push(() -> {
                            gammaControlPoints.remove(addedPoint);
                            if (selectedGammaControlPoint == addedPoint) {
                                selectedGammaControlPoint = null;
                            }
                            if (hoveredGammaControlPoint == addedPoint) {
                                hoveredGammaControlPoint = null;
                            }
                            renumberGammaControlPoints();
                            repaint();
                        });
                        renumberGammaControlPoints();
                        repaint();
                        return;
                    }
                    if (gammaProfileMode) {
                        GammaProfile profile = findGammaProfile(e.getPoint());
                        if (profile != null) {
                            activeGammaProfile = profile;
                            selectedGammaProfile = profile;
                            selectedNoisePoint = null;
                            selectedGammaControlPoint = null;
                            selectedPprPoint = null;
                            repaint();
                            return;
                        }
                        if (gammaProfiles.size() >= 10) {
                            Toolkit.getDefaultToolkit().beep();
                            return;
                        }
                        GammaProfile newProfile = createGammaProfileFromLineEnd(e.getPoint());
                        if (newProfile == null) {
                            return;
                        }
                        gammaProfiles.add(newProfile);
                        selectedGammaProfile = newProfile;
                        selectedNoisePoint = null;
                        selectedGammaControlPoint = null;
                        selectedPprPoint = null;
                        if (!editGammaProfileNumber(newProfile)) {
                            gammaProfiles.remove(newProfile);
                            selectedGammaProfile = null;
                            repaint();
                            return;
                        }
                        activeGammaProfile = newProfile;
                        currentMousePoint = null;
                        undoActions.push(() -> {
                            gammaProfiles.remove(newProfile);
                            if (selectedGammaProfile == newProfile) {
                                selectedGammaProfile = null;
                            }
                            if (activeGammaProfile == newProfile) {
                                activeGammaProfile = null;
                            }
                            fireGammaProfilesChanged();
                            repaint();
                        });
                        fireGammaProfilesChanged();
                        repaint();
                        return;
                    }
                    if (!boundaryMode) {
                        return;
                    }
                    Point2D.Double imagePoint = toBoundaryImagePoint(e.getPoint());
                    if (imagePoint == null) {
                        return;
                    }
                    int addedIndex = boundaryPoints.size();
                    boundaryPoints.add(imagePoint);
                    undoActions.push(() -> {
                        if (addedIndex < boundaryPoints.size() && boundaryPoints.get(addedIndex) == imagePoint) {
                            boundaryPoints.remove(addedIndex);
                        } else {
                            boundaryPoints.remove(imagePoint);
                        }
                        repaint();
                    });
                    currentMousePoint = snapBoundaryPreviewPoint(e.getPoint());
                    repaint();
                }

                @Override
                public void mouseReleased(MouseEvent e) {
                    activeGammaProfile = null;
                    activeGammaControlPoint = null;
                    activePprPoint = null;
                    activeNoisePoint = null;
                }

                @Override
                public void mouseClicked(MouseEvent e) {
                    if (!gammaProfileMode || e.getClickCount() != 2 || image == null) {
                        return;
                    }
                    GammaProfile profile = findGammaProfile(e.getPoint());
                    if (profile != null) {
                        selectedGammaProfile = profile;
                        editGammaProfileNumber(profile);
                    }
                }

                @Override
                public void mouseMoved(MouseEvent e) {
                    GammaControlPoint point = findGammaControlPoint(e.getPoint());
                    PprPoint pprPoint = findPprPoint(e.getPoint());
                    NoisePoint noisePoint = findNoisePoint(e.getPoint());
                    if (point != null || pprPoint != null || noisePoint != null) {
                        requestFocusInWindow();
                    }
                    if (point != hoveredGammaControlPoint) {
                        hoveredGammaControlPoint = point;
                        repaint();
                    }
                    if (pprPoint != hoveredPprPoint) {
                        hoveredPprPoint = pprPoint;
                        repaint();
                    }
                    if (noisePoint != hoveredNoisePoint) {
                        hoveredNoisePoint = noisePoint;
                        repaint();
                    }
                    if (!boundaryMode) {
                        return;
                    }
                    currentMousePoint = snapBoundaryPreviewPoint(e.getPoint());
                    repaint();
                }

                @Override
                public void mouseExited(MouseEvent e) {
                    if (hoveredGammaControlPoint != null || hoveredPprPoint != null
                            || hoveredNoisePoint != null || currentMousePoint != null) {
                        hoveredGammaControlPoint = null;
                        hoveredPprPoint = null;
                        hoveredNoisePoint = null;
                        currentMousePoint = null;
                        repaint();
                    }
                }

                @Override
                public void mouseDragged(MouseEvent e) {
                    if ((e.getModifiersEx() & InputEvent.BUTTON1_DOWN_MASK) == 0) {
                        mouseMoved(e);
                        return;
                    }
                    if (noisePointMode && activeNoisePoint != null) {
                        moveNoisePointTo(activeNoisePoint, e.getPoint());
                        return;
                    }
                    if (pprPointMode && activePprPoint != null) {
                        movePprPointTo(activePprPoint, e.getPoint());
                        return;
                    }
                    if (gammaControlPointMode && activeGammaControlPoint != null) {
                        moveGammaControlPointTo(activeGammaControlPoint, e.getPoint());
                        return;
                    }
                    if (gammaProfileMode && activeGammaProfile != null) {
                        updateGammaProfileAngle(activeGammaProfile, e.getPoint());
                        repaint();
                        return;
                    }
                    mouseMoved(e);
                }
            };
            addMouseListener(mouseHandler);
            addMouseMotionListener(mouseHandler);

            getInputMap(WHEN_FOCUSED).put(
                    KeyStroke.getKeyStroke(KeyEvent.VK_Z, InputEvent.CTRL_DOWN_MASK),
                    "undoBoundaryPoint"
            );
            getActionMap().put("undoBoundaryPoint", new AbstractAction() {
                @Override
                public void actionPerformed(java.awt.event.ActionEvent e) {
                    undoLastAction();
                }
            });
            getInputMap(WHEN_FOCUSED).put(
                    KeyStroke.getKeyStroke(KeyEvent.VK_DELETE, 0),
                    "deleteGammaProfile"
            );
            getActionMap().put("deleteGammaProfile", new AbstractAction() {
                @Override
                public void actionPerformed(java.awt.event.ActionEvent e) {
                    deleteSelectedSketchItem();
                }
            });
            bindMoveKey(KeyEvent.VK_LEFT, -1, 0);
            bindMoveKey(KeyEvent.VK_RIGHT, 1, 0);
            bindMoveKey(KeyEvent.VK_UP, 0, -1);
            bindMoveKey(KeyEvent.VK_DOWN, 0, 1);
        }

        void setImage(BufferedImage image) {
            this.image = image;
            boundaryPoints.clear();
            gammaProfiles.clear();
            gammaControlPoints.clear();
            pprPoints.clear();
            noisePoints.clear();
            activeGammaProfile = null;
            activeGammaControlPoint = null;
            activePprPoint = null;
            activeNoisePoint = null;
            selectedGammaProfile = null;
            selectedGammaControlPoint = null;
            hoveredGammaControlPoint = null;
            selectedPprPoint = null;
            hoveredPprPoint = null;
            selectedNoisePoint = null;
            hoveredNoisePoint = null;
            undoActions.clear();
            currentMousePoint = null;
            fireGammaProfilesChanged();
            repaint();
        }

        void setBoundaryMode(boolean boundaryMode) {
            this.boundaryMode = boundaryMode;
            if (boundaryMode) {
                gammaProfileMode = false;
                gammaControlPointMode = false;
                pprPointMode = false;
                noisePointMode = false;
            }
            currentMousePoint = null;
            updateCursor();
            repaint();
        }

        void setGammaProfileMode(boolean gammaProfileMode) {
            this.gammaProfileMode = gammaProfileMode;
            if (gammaProfileMode) {
                boundaryMode = false;
                gammaControlPointMode = false;
                pprPointMode = false;
                noisePointMode = false;
            }
            currentMousePoint = null;
            activeGammaProfile = null;
            updateCursor();
            repaint();
        }

        void setGammaControlPointMode(boolean gammaControlPointMode) {
            this.gammaControlPointMode = gammaControlPointMode;
            if (gammaControlPointMode) {
                boundaryMode = false;
                gammaProfileMode = false;
                pprPointMode = false;
                noisePointMode = false;
            }
            currentMousePoint = null;
            activeGammaProfile = null;
            updateCursor();
            repaint();
        }

        void setPprPointMode(boolean pprPointMode) {
            this.pprPointMode = pprPointMode;
            if (pprPointMode) {
                boundaryMode = false;
                gammaProfileMode = false;
                gammaControlPointMode = false;
                noisePointMode = false;
            }
            currentMousePoint = null;
            activeGammaProfile = null;
            updateCursor();
            repaint();
        }

        void setNoisePointMode(boolean noisePointMode) {
            this.noisePointMode = noisePointMode;
            if (noisePointMode) {
                boundaryMode = false;
                gammaProfileMode = false;
                gammaControlPointMode = false;
                pprPointMode = false;
            }
            currentMousePoint = null;
            activeGammaProfile = null;
            updateCursor();
            repaint();
        }

        void setRadiationLayerVisible(boolean radiationLayerVisible) {
            this.radiationLayerVisible = radiationLayerVisible;
            repaint();
        }

        void setNoiseLayerVisible(boolean noiseLayerVisible) {
            this.noiseLayerVisible = noiseLayerVisible;
            repaint();
        }

        void setBoundaryColor(Color boundaryColor) {
            this.boundaryColor = boundaryColor == null ? Color.BLACK : boundaryColor;
            repaint();
        }

        void setGammaControlPointScale(int gammaControlPointScale) {
            this.gammaControlPointScale = Math.max(60, Math.min(160, gammaControlPointScale));
            repaint();
        }

        int getGammaControlPointScale() {
            return gammaControlPointScale;
        }

        void setPprPointScale(int pprPointScale) {
            this.pprPointScale = Math.max(60, Math.min(160, pprPointScale));
            repaint();
        }

        int getPprPointScale() {
            return pprPointScale;
        }

        void setNoisePointScale(int noisePointScale) {
            this.noisePointScale = Math.max(60, Math.min(160, noisePointScale));
            repaint();
        }

        int getNoisePointScale() {
            return noisePointScale;
        }

        BufferedImage getImage() {
            return image;
        }

        Color getBoundaryColor() {
            return boundaryColor;
        }

        List<PointSnapshot> getBoundaryPointSnapshots() {
            List<PointSnapshot> snapshots = new ArrayList<>();
            for (Point2D.Double point : boundaryPoints) {
                snapshots.add(new PointSnapshot(point.x, point.y));
            }
            return snapshots;
        }

        List<GammaProfileSnapshot> getGammaProfileSnapshots() {
            List<GammaProfileSnapshot> snapshots = new ArrayList<>();
            for (GammaProfile profile : gammaProfiles) {
                snapshots.add(new GammaProfileSnapshot(profile.x, profile.y, profile.angle, profile.number));
            }
            return snapshots;
        }

        int getMaxGammaProfileNumber() {
            int maxNumber = 0;
            for (GammaProfile profile : gammaProfiles) {
                if (profile.number == null || profile.number.isBlank()) {
                    continue;
                }
                try {
                    maxNumber = Math.max(maxNumber, Integer.parseInt(profile.number.trim()));
                } catch (NumberFormatException ignored) {
                }
            }
            return maxNumber;
        }

        int getGammaControlPointCount() {
            return gammaControlPoints.size();
        }

        int getPprPointCount() {
            return pprPoints.size();
        }

        int getNoisePointCount() {
            return noisePoints.size();
        }

        List<GammaControlPointSnapshot> getGammaControlPointSnapshots() {
            List<GammaControlPointSnapshot> snapshots = new ArrayList<>();
            for (GammaControlPoint point : gammaControlPoints) {
                snapshots.add(new GammaControlPointSnapshot(point.x, point.y, point.number));
            }
            return snapshots;
        }

        List<PprPointSnapshot> getPprPointSnapshots() {
            List<PprPointSnapshot> snapshots = new ArrayList<>();
            for (PprPoint point : pprPoints) {
                snapshots.add(new PprPointSnapshot(point.x, point.y, point.number));
            }
            return snapshots;
        }

        List<NoisePointSnapshot> getNoisePointSnapshots() {
            List<NoisePointSnapshot> snapshots = new ArrayList<>();
            for (NoisePoint point : noisePoints) {
                snapshots.add(new NoisePointSnapshot(point.x, point.y, point.number));
            }
            return snapshots;
        }

        void applySketchSnapshot(AreaProjectSnapshot snapshot, BufferedImage loadedImage) {
            image = loadedImage;
            boundaryPoints.clear();
            gammaProfiles.clear();
            gammaControlPoints.clear();
            pprPoints.clear();
            noisePoints.clear();
            if (snapshot.boundaryPoints != null) {
                for (PointSnapshot point : snapshot.boundaryPoints) {
                    boundaryPoints.add(new Point2D.Double(point.x, point.y));
                }
            }
            if (snapshot.gammaProfiles != null) {
                for (GammaProfileSnapshot profile : snapshot.gammaProfiles) {
                    GammaProfile gammaProfile = new GammaProfile(profile.x, profile.y, profile.angle);
                    gammaProfile.number = profile.number == null ? "" : profile.number;
                    gammaProfiles.add(gammaProfile);
                }
            }
            if (snapshot.gammaControlPoints != null) {
                for (GammaControlPointSnapshot point : snapshot.gammaControlPoints) {
                    gammaControlPoints.add(new GammaControlPoint(point.x, point.y, point.number));
                }
            }
            if (snapshot.pprPoints != null) {
                for (PprPointSnapshot point : snapshot.pprPoints) {
                    pprPoints.add(new PprPoint(point.x, point.y, point.number));
                }
            }
            if (snapshot.noisePoints != null) {
                for (NoisePointSnapshot point : snapshot.noisePoints) {
                    noisePoints.add(new NoisePoint(point.x, point.y, point.number));
                }
            }
            boundaryColor = snapshot.boundaryColorRgb == 0
                    ? Color.BLACK
                    : new Color(snapshot.boundaryColorRgb, true);
            gammaControlPointScale = Math.max(60, Math.min(160,
                    snapshot.gammaControlPointScale <= 0 ? 100 : snapshot.gammaControlPointScale));
            pprPointScale = Math.max(60, Math.min(160,
                    snapshot.pprPointScale <= 0 ? 100 : snapshot.pprPointScale));
            noisePointScale = Math.max(60, Math.min(160,
                    snapshot.noisePointScale <= 0 ? 100 : snapshot.noisePointScale));
            activeGammaProfile = null;
            activeGammaControlPoint = null;
            activePprPoint = null;
            activeNoisePoint = null;
            selectedGammaProfile = null;
            selectedGammaControlPoint = null;
            hoveredGammaControlPoint = null;
            selectedPprPoint = null;
            hoveredPprPoint = null;
            selectedNoisePoint = null;
            hoveredNoisePoint = null;
            currentMousePoint = null;
            undoActions.clear();
            renumberGammaControlPoints();
            renumberPprPoints();
            renumberNoisePoints();
            fireGammaProfilesChanged();
            repaint();
        }

        void setBoundaryModeOffAction(Runnable boundaryModeOffAction) {
            this.boundaryModeOffAction = boundaryModeOffAction;
        }

        void setGammaProfilesChangedAction(Runnable gammaProfilesChangedAction) {
            this.gammaProfilesChangedAction = gammaProfilesChangedAction == null ? () -> {
            } : gammaProfilesChangedAction;
        }

        private void fireGammaProfilesChanged() {
            gammaProfilesChangedAction.run();
        }

        void undoLastAction() {
            if (undoActions.isEmpty()) {
                return;
            }
            undoActions.pop().run();
        }

        void deleteSelectedSketchItem() {
            if (hoveredNoisePoint != null || selectedNoisePoint != null) {
                NoisePoint point = hoveredNoisePoint != null ? hoveredNoisePoint : selectedNoisePoint;
                int removedIndex = noisePoints.indexOf(point);
                if (removedIndex < 0) {
                    return;
                }
                noisePoints.remove(removedIndex);
                if (point == selectedNoisePoint) {
                    selectedNoisePoint = null;
                }
                hoveredNoisePoint = null;
                undoActions.push(() -> {
                    int insertIndex = Math.min(removedIndex, noisePoints.size());
                    noisePoints.add(insertIndex, point);
                    selectedNoisePoint = point;
                    hoveredNoisePoint = point;
                    renumberNoisePoints();
                    repaint();
                });
                renumberNoisePoints();
                repaint();
                return;
            }
            if (hoveredPprPoint != null || selectedPprPoint != null) {
                PprPoint point = hoveredPprPoint != null ? hoveredPprPoint : selectedPprPoint;
                int removedIndex = pprPoints.indexOf(point);
                if (removedIndex < 0) {
                    return;
                }
                pprPoints.remove(removedIndex);
                if (point == selectedPprPoint) {
                    selectedPprPoint = null;
                }
                hoveredPprPoint = null;
                undoActions.push(() -> {
                    int insertIndex = Math.min(removedIndex, pprPoints.size());
                    pprPoints.add(insertIndex, point);
                    selectedPprPoint = point;
                    hoveredPprPoint = point;
                    renumberPprPoints();
                    repaint();
                });
                renumberPprPoints();
                repaint();
                return;
            }
            if (hoveredGammaControlPoint != null || selectedGammaControlPoint != null) {
                GammaControlPoint point = hoveredGammaControlPoint != null
                        ? hoveredGammaControlPoint
                        : selectedGammaControlPoint;
                int removedIndex = gammaControlPoints.indexOf(point);
                if (removedIndex < 0) {
                    return;
                }
                gammaControlPoints.remove(removedIndex);
                if (point == selectedGammaControlPoint) {
                    selectedGammaControlPoint = null;
                }
                hoveredGammaControlPoint = null;
                undoActions.push(() -> {
                    int insertIndex = Math.min(removedIndex, gammaControlPoints.size());
                    gammaControlPoints.add(insertIndex, point);
                    selectedGammaControlPoint = point;
                    hoveredGammaControlPoint = point;
                    renumberGammaControlPoints();
                    repaint();
                });
                renumberGammaControlPoints();
                repaint();
                return;
            }
            deleteSelectedGammaProfile();
        }

        void deleteSelectedGammaProfile() {
            if (selectedGammaProfile == null) {
                return;
            }
            GammaProfile removedProfile = selectedGammaProfile;
            int removedIndex = gammaProfiles.indexOf(removedProfile);
            if (removedIndex < 0) {
                return;
            }
            gammaProfiles.remove(removedIndex);
            selectedGammaProfile = null;
            activeGammaProfile = null;
            undoActions.push(() -> {
                int insertIndex = Math.min(removedIndex, gammaProfiles.size());
                gammaProfiles.add(insertIndex, removedProfile);
                selectedGammaProfile = removedProfile;
                fireGammaProfilesChanged();
                repaint();
            });
            fireGammaProfilesChanged();
            repaint();
        }

        private void bindMoveKey(int keyCode, int dx, int dy) {
            String actionName = "moveGammaProfile" + keyCode;
            getInputMap(WHEN_FOCUSED).put(KeyStroke.getKeyStroke(keyCode, 0), actionName);
            getActionMap().put(actionName, new AbstractAction() {
                @Override
                public void actionPerformed(java.awt.event.ActionEvent e) {
                    moveSelectedSketchItem(dx, dy);
                }
            });
        }

        private void moveSelectedSketchItem(int dx, int dy) {
            if (selectedNoisePoint != null) {
                moveSelectedNoisePoint(dx, dy);
                return;
            }
            if (selectedPprPoint != null) {
                moveSelectedPprPoint(dx, dy);
                return;
            }
            if (selectedGammaControlPoint != null) {
                moveSelectedGammaControlPoint(dx, dy);
                return;
            }
            moveSelectedGammaProfile(dx, dy);
        }

        private void moveSelectedNoisePoint(int dx, int dy) {
            if (image == null) {
                return;
            }
            updateImageBounds();
            Point center = toScreenPoint(new Point2D.Double(selectedNoisePoint.x, selectedNoisePoint.y));
            Point2D.Double movedPoint = toImagePoint(new Point(center.x + dx, center.y + dy));
            if (movedPoint == null || !isInsideBoundary(movedPoint)) {
                return;
            }
            selectedNoisePoint.x = movedPoint.x;
            selectedNoisePoint.y = movedPoint.y;
            repaint();
        }

        private void moveNoisePointTo(NoisePoint point, Point screenPoint) {
            if (image == null) {
                return;
            }
            Point2D.Double movedPoint = toImagePoint(screenPoint);
            if (movedPoint == null || !isInsideBoundary(movedPoint)) {
                return;
            }
            point.x = movedPoint.x;
            point.y = movedPoint.y;
            selectedNoisePoint = point;
            hoveredNoisePoint = point;
            repaint();
        }

        private void moveSelectedPprPoint(int dx, int dy) {
            if (image == null) {
                return;
            }
            updateImageBounds();
            Point center = toScreenPoint(new Point2D.Double(selectedPprPoint.x, selectedPprPoint.y));
            Point2D.Double movedPoint = toImagePoint(new Point(center.x + dx, center.y + dy));
            if (movedPoint == null || !isInsideBoundary(movedPoint) || !isFarFromGammaMarks(movedPoint)) {
                return;
            }
            selectedPprPoint.x = movedPoint.x;
            selectedPprPoint.y = movedPoint.y;
            repaint();
        }

        private void movePprPointTo(PprPoint point, Point screenPoint) {
            if (image == null) {
                return;
            }
            Point2D.Double movedPoint = toImagePoint(screenPoint);
            if (movedPoint == null || !isInsideBoundary(movedPoint) || !isFarFromGammaMarks(movedPoint)) {
                return;
            }
            point.x = movedPoint.x;
            point.y = movedPoint.y;
            selectedPprPoint = point;
            hoveredPprPoint = point;
            repaint();
        }

        private void moveSelectedGammaControlPoint(int dx, int dy) {
            if (image == null) {
                return;
            }
            updateImageBounds();
            Point center = toScreenPoint(new Point2D.Double(selectedGammaControlPoint.x, selectedGammaControlPoint.y));
            Point2D.Double movedPoint = toImagePoint(new Point(center.x + dx, center.y + dy));
            if (movedPoint == null) {
                return;
            }
            selectedGammaControlPoint.x = movedPoint.x;
            selectedGammaControlPoint.y = movedPoint.y;
            repaint();
        }

        private void moveGammaControlPointTo(GammaControlPoint point, Point screenPoint) {
            if (image == null) {
                return;
            }
            Point2D.Double movedPoint = toImagePoint(screenPoint);
            if (movedPoint == null) {
                return;
            }
            point.x = movedPoint.x;
            point.y = movedPoint.y;
            selectedGammaControlPoint = point;
            hoveredGammaControlPoint = point;
            repaint();
        }

        private void moveSelectedGammaProfile(int dx, int dy) {
            if (selectedGammaProfile == null || image == null) {
                return;
            }
            updateImageBounds();
            Point center = toScreenPoint(new Point2D.Double(selectedGammaProfile.x, selectedGammaProfile.y));
            Point2D.Double movedPoint = toImagePoint(new Point(center.x + dx, center.y + dy));
            if (movedPoint == null) {
                return;
            }
            selectedGammaProfile.x = movedPoint.x;
            selectedGammaProfile.y = movedPoint.y;
            repaint();
        }

        private void updateCursor() {
            setCursor((boundaryMode || gammaProfileMode || gammaControlPointMode || pprPointMode || noisePointMode)
                    ? Cursor.getPredefinedCursor(Cursor.CROSSHAIR_CURSOR)
                    : Cursor.getDefaultCursor());
        }

        @Override
        protected void paintComponent(Graphics g) {
            super.paintComponent(g);
            Graphics2D g2 = (Graphics2D) g.create();
            try {
                g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
                if (image == null) {
                    Color disabledColor = UIManager.getColor("Label.disabledForeground");
                    g2.setColor(disabledColor == null ? Color.GRAY : disabledColor);
                    String text = "Эскиз не выбран";
                    FontMetrics metrics = g2.getFontMetrics();
                    int x = (getWidth() - metrics.stringWidth(text)) / 2;
                    int y = (getHeight() + metrics.getAscent()) / 2;
                    g2.drawString(text, Math.max(12, x), Math.max(24, y));
                    return;
                }
                updateImageBounds();
                g2.drawImage(image, imageBounds.x, imageBounds.y, imageBounds.width, imageBounds.height, null);
                if (radiationLayerVisible || noiseLayerVisible || boundaryMode) {
                    paintBoundary(g2);
                }
                if (radiationLayerVisible) {
                    paintGammaControlPoints(g2);
                    paintPprPoints(g2);
                    paintGammaProfiles(g2);
                }
                if (noiseLayerVisible) {
                    paintNoisePoints(g2);
                }
            } finally {
                g2.dispose();
            }
        }

        BufferedImage renderSketchImage(boolean includeNoisePoints) {
            if (image == null) {
                return null;
            }
            Dimension renderSize = resolveSketchRenderSize();
            BufferedImage renderedImage = new BufferedImage(
                    renderSize.width,
                    renderSize.height,
                    BufferedImage.TYPE_INT_RGB
            );
            Graphics2D g2 = renderedImage.createGraphics();
            Rectangle previousBounds = new Rectangle(imageBounds);
            Point previousMousePoint = currentMousePoint;
            try {
                g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g2.setColor(Color.WHITE);
                g2.fillRect(0, 0, renderSize.width, renderSize.height);
                imageBounds.setBounds(0, 0, renderSize.width, renderSize.height);
                currentMousePoint = null;
                g2.drawImage(image, 0, 0, renderSize.width, renderSize.height, null);
                paintBoundary(g2);
                paintGammaControlPoints(g2);
                paintPprPoints(g2);
                if (includeNoisePoints) {
                    paintNoisePoints(g2);
                }
                paintGammaProfiles(g2);
            } finally {
                currentMousePoint = previousMousePoint;
                imageBounds.setBounds(previousBounds);
                g2.dispose();
            }
            return renderedImage;
        }

        BufferedImage renderNoiseSketchImage() {
            if (image == null) {
                return null;
            }
            Dimension renderSize = resolveSketchRenderSize();
            BufferedImage renderedImage = new BufferedImage(
                    renderSize.width,
                    renderSize.height,
                    BufferedImage.TYPE_INT_RGB
            );
            Graphics2D g2 = renderedImage.createGraphics();
            Rectangle previousBounds = new Rectangle(imageBounds);
            Point previousMousePoint = currentMousePoint;
            try {
                g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g2.setColor(Color.WHITE);
                g2.fillRect(0, 0, renderSize.width, renderSize.height);
                imageBounds.setBounds(0, 0, renderSize.width, renderSize.height);
                currentMousePoint = null;
                g2.drawImage(image, 0, 0, renderSize.width, renderSize.height, null);
                paintBoundary(g2);
                paintNoisePoints(g2);
            } finally {
                currentMousePoint = previousMousePoint;
                imageBounds.setBounds(previousBounds);
                g2.dispose();
            }
            return renderedImage;
        }

        private Dimension resolveSketchRenderSize() {
            if (imageBounds.width > 0 && imageBounds.height > 0) {
                return new Dimension(imageBounds.width, imageBounds.height);
            }
            int availableWidth = Math.max(900, getWidth() - 48);
            int availableHeight = Math.max(560, getHeight() - 48);
            double scale = Math.min(
                    availableWidth / (double) image.getWidth(),
                    availableHeight / (double) image.getHeight()
            );
            scale = Math.max(0.05d, scale);
            return new Dimension(
                    Math.max(1, (int) Math.round(image.getWidth() * scale)),
                    Math.max(1, (int) Math.round(image.getHeight() * scale))
            );
        }

        private void updateImageBounds() {
            int padding = 24;
            double scaleX = (getWidth() - padding * 2) / (double) image.getWidth();
            double scaleY = (getHeight() - padding * 2) / (double) image.getHeight();
            double scale = Math.max(0.05d, Math.min(scaleX, scaleY));
            int width = Math.max(1, (int) Math.round(image.getWidth() * scale));
            int height = Math.max(1, (int) Math.round(image.getHeight() * scale));
            int x = (getWidth() - width) / 2;
            int y = (getHeight() - height) / 2;
            imageBounds.setBounds(x, y, width, height);
        }

        private void paintBoundary(Graphics2D g2) {
            g2.setColor(boundaryColor);
            g2.setStroke(new BasicStroke(
                    3f,
                    BasicStroke.CAP_BUTT,
                    BasicStroke.JOIN_MITER,
                    10f,
                    new float[]{8f, 6f},
                    0f
            ));
            for (int i = 1; i < boundaryPoints.size(); i++) {
                Point p1 = toScreenPoint(boundaryPoints.get(i - 1));
                Point p2 = toScreenPoint(boundaryPoints.get(i));
                g2.drawLine(p1.x, p1.y, p2.x, p2.y);
            }
            if (boundaryMode && !boundaryPoints.isEmpty() && currentMousePoint != null
                    && imageBounds.contains(currentMousePoint)) {
                Point last = toScreenPoint(boundaryPoints.get(boundaryPoints.size() - 1));
                g2.drawLine(last.x, last.y, currentMousePoint.x, currentMousePoint.y);
            }
        }

        private void paintGammaProfiles(Graphics2D g2) {
            Font font = new Font("Arial", Font.PLAIN, 16);
            g2.setFont(font);
            for (GammaProfile profile : gammaProfiles) {
                Point center = toScreenPoint(new Point2D.Double(profile.x, profile.y));
                int lineStartX = center.x + (int) Math.round(Math.cos(profile.angle) * GAMMA_RADIUS);
                int lineStartY = center.y + (int) Math.round(Math.sin(profile.angle) * GAMMA_RADIUS);
                int lineEndX = center.x + (int) Math.round(Math.cos(profile.angle) * (GAMMA_RADIUS + GAMMA_LINE_LENGTH));
                int lineEndY = center.y + (int) Math.round(Math.sin(profile.angle) * (GAMMA_RADIUS + GAMMA_LINE_LENGTH));

                g2.setColor(Color.BLACK);
                g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER));
                g2.drawLine(lineStartX, lineStartY, lineEndX, lineEndY);

                g2.setColor(Color.WHITE);
                g2.fillOval(center.x - GAMMA_RADIUS, center.y - GAMMA_RADIUS, GAMMA_RADIUS * 2, GAMMA_RADIUS * 2);
                g2.setColor(Color.BLACK);
                g2.setStroke(new BasicStroke(profile == selectedGammaProfile ? 1.8f : 1.2f));
                g2.drawOval(center.x - GAMMA_RADIUS, center.y - GAMMA_RADIUS, GAMMA_RADIUS * 2, GAMMA_RADIUS * 2);
                g2.setStroke(new BasicStroke(1.2f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER));

                String numberText = profile.number == null ? "" : profile.number;
                FontMetrics metrics = g2.getFontMetrics(font);
                int textX = center.x - metrics.stringWidth(numberText) / 2;
                int textY = center.y + (metrics.getAscent() - metrics.getDescent()) / 2;
                g2.drawString(numberText, textX, textY);
            }
        }

        void insertGammaControlPoints(int count) {
            if (image == null) {
                JOptionPane.showMessageDialog(this,
                        "Сначала добавьте эскиз.",
                        "КТ гамма",
                        JOptionPane.WARNING_MESSAGE);
                return;
            }
            if (boundaryPoints.size() < 3) {
                JOptionPane.showMessageDialog(this,
                        "Сначала нарисуйте границу участка.",
                        "КТ гамма",
                        JOptionPane.WARNING_MESSAGE);
                return;
            }
            List<GammaControlPoint> previousPoints = new ArrayList<>(gammaControlPoints);
            GammaControlPoint previousSelectedPoint = selectedGammaControlPoint;
            GammaControlPoint previousHoveredPoint = hoveredGammaControlPoint;
            gammaControlPoints.clear();
            List<Point2D.Double> points = generateGammaControlPoints(count);
            for (int i = 0; i < points.size(); i++) {
                Point2D.Double point = points.get(i);
                gammaControlPoints.add(new GammaControlPoint(point.x, point.y, ""));
            }
            selectedGammaControlPoint = null;
            hoveredGammaControlPoint = null;
            undoActions.push(() -> {
                gammaControlPoints.clear();
                gammaControlPoints.addAll(previousPoints);
                selectedGammaControlPoint = previousPoints.contains(previousSelectedPoint) ? previousSelectedPoint : null;
                hoveredGammaControlPoint = previousPoints.contains(previousHoveredPoint) ? previousHoveredPoint : null;
                renumberGammaControlPoints();
                repaint();
            });
            renumberGammaControlPoints();
            repaint();
        }

        void insertPprPoints(int count) {
            if (image == null) {
                JOptionPane.showMessageDialog(this,
                        "Сначала добавьте эскиз.",
                        "ППР",
                        JOptionPane.WARNING_MESSAGE);
                return;
            }
            if (boundaryPoints.size() < 3) {
                JOptionPane.showMessageDialog(this,
                        "Сначала нарисуйте границу участка.",
                        "ППР",
                        JOptionPane.WARNING_MESSAGE);
                return;
            }
            List<PprPoint> previousPoints = new ArrayList<>(pprPoints);
            PprPoint previousSelectedPoint = selectedPprPoint;
            PprPoint previousHoveredPoint = hoveredPprPoint;
            pprPoints.clear();
            List<Point2D.Double> points = generatePprPoints(count);
            for (Point2D.Double point : points) {
                pprPoints.add(new PprPoint(point.x, point.y, ""));
            }
            selectedPprPoint = null;
            hoveredPprPoint = null;
            undoActions.push(() -> {
                pprPoints.clear();
                pprPoints.addAll(previousPoints);
                selectedPprPoint = previousPoints.contains(previousSelectedPoint) ? previousSelectedPoint : null;
                hoveredPprPoint = previousPoints.contains(previousHoveredPoint) ? previousHoveredPoint : null;
                renumberPprPoints();
                repaint();
            });
            renumberPprPoints();
            repaint();
            if (points.size() < count) {
                JOptionPane.showMessageDialog(this,
                        "Удалось разместить только " + points.size() + " из " + count + " точек ППР.",
                        "ППР",
                        JOptionPane.WARNING_MESSAGE);
            }
        }

        private void paintGammaControlPoints(Graphics2D g2) {
            Font font = new Font("Arial", Font.PLAIN, 16);
            g2.setFont(font);
            g2.setColor(Color.BLACK);
            g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER));
            int radius = getGammaControlPointRadius();
            int size = radius * 2;
            for (GammaControlPoint point : gammaControlPoints) {
                Point center = toScreenPoint(new Point2D.Double(point.x, point.y));
                int x = center.x - radius;
                int y = center.y - radius;
                g2.setColor(Color.WHITE);
                g2.fillRect(x, y, size, size);
                g2.setColor(Color.BLACK);
                g2.setStroke(new BasicStroke(
                        point == selectedGammaControlPoint || point == hoveredGammaControlPoint ? 4f : 3f,
                        BasicStroke.CAP_BUTT,
                        BasicStroke.JOIN_MITER
                ));
                g2.drawRect(x, y, size, size);
                String numberText = point.number == null ? "" : point.number;
                FontMetrics metrics = g2.getFontMetrics(font);
                int textX = center.x - metrics.stringWidth(numberText) / 2;
                int textY = center.y + (metrics.getAscent() - metrics.getDescent()) / 2;
                g2.drawString(numberText, textX, textY);
            }
        }

        private void paintPprPoints(Graphics2D g2) {
            Font font = new Font("Arial", Font.PLAIN, 16);
            g2.setFont(font);
            int radius = getPprPointRadius();
            for (PprPoint point : pprPoints) {
                Point center = toScreenPoint(new Point2D.Double(point.x, point.y));
                Polygon diamond = new Polygon(
                        new int[]{center.x, center.x + radius, center.x, center.x - radius},
                        new int[]{center.y - radius, center.y, center.y + radius, center.y},
                        4
                );
                g2.setColor(Color.WHITE);
                g2.fillPolygon(diamond);
                g2.setColor(Color.BLACK);
                g2.setStroke(new BasicStroke(
                        point == selectedPprPoint || point == hoveredPprPoint ? 4f : 3f,
                        BasicStroke.CAP_BUTT,
                        BasicStroke.JOIN_MITER
                ));
                g2.drawPolygon(diamond);
                String numberText = point.number == null ? "" : point.number;
                FontMetrics metrics = g2.getFontMetrics(font);
                int textX = center.x - metrics.stringWidth(numberText) / 2;
                int textY = center.y + (metrics.getAscent() - metrics.getDescent()) / 2;
                g2.drawString(numberText, textX, textY);
            }
        }

        private void paintNoisePoints(Graphics2D g2) {
            Font font = new Font("Arial", Font.PLAIN, 24);
            g2.setFont(font);
            int radius = getNoisePointRadius();
            int size = radius * 2;
            for (NoisePoint point : noisePoints) {
                Point center = toScreenPoint(new Point2D.Double(point.x, point.y));
                int x = center.x - radius;
                int y = center.y - radius;
                g2.setColor(Color.BLACK);
                g2.fillOval(x, y, size, size);
                if (point == selectedNoisePoint || point == hoveredNoisePoint) {
                    g2.setColor(Color.WHITE);
                    g2.setStroke(new BasicStroke(2f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                    g2.drawOval(x + 2, y + 2, Math.max(1, size - 4), Math.max(1, size - 4));
                }
                String numberText = point.number == null ? "" : point.number;
                FontMetrics metrics = g2.getFontMetrics(font);
                int textX = center.x + radius + 4;
                int textY = center.y + (metrics.getAscent() - metrics.getDescent()) / 2;
                g2.setColor(Color.BLACK);
                g2.drawString("T" + numberText, textX, textY);
            }
        }

        private List<Point2D.Double> generateGammaControlPoints(int count) {
            Rectangle2D.Double bounds = boundaryBounds();
            if (bounds.width <= 0d || bounds.height <= 0d) {
                return List.of();
            }
            List<Point2D.Double> candidates = new ArrayList<>();
            double ratio = bounds.width / bounds.height;
            int baseColumns = Math.max(1, (int) Math.ceil(Math.sqrt(count * ratio)));
            int baseRows = Math.max(1, (int) Math.ceil(count / (double) baseColumns));

            for (int multiplier = 1; multiplier <= 8 && candidates.size() < count; multiplier++) {
                candidates.clear();
                int columns = baseColumns * multiplier;
                int rows = baseRows * multiplier;
                for (int row = 0; row < rows; row++) {
                    double y = bounds.y + bounds.height * (row + 0.5d) / rows;
                    for (int col = 0; col < columns; col++) {
                        double x = bounds.x + bounds.width * (col + 0.5d) / columns;
                        Point2D.Double point = new Point2D.Double(x, y);
                        if (isInsideBoundary(point)) {
                            candidates.add(point);
                        }
                    }
                }
            }

            if (candidates.isEmpty()) {
                return List.of();
            }
            if (candidates.size() <= count) {
                return new ArrayList<>(candidates);
            }
            List<Point2D.Double> selected = new ArrayList<>();
            for (int i = 0; i < count; i++) {
                int index = count == 1
                        ? candidates.size() / 2
                        : (int) Math.round(i * (candidates.size() - 1) / (double) (count - 1));
                selected.add(candidates.get(index));
            }
            return selected;
        }

        private List<Point2D.Double> generatePprPoints(int count) {
            Rectangle2D.Double bounds = boundaryBounds();
            if (bounds.width <= 0d || bounds.height <= 0d) {
                return List.of();
            }
            updateImageBounds();
            List<Point2D.Double> candidates = new ArrayList<>();
            double ratio = bounds.width / bounds.height;
            int baseColumns = Math.max(1, (int) Math.ceil(Math.sqrt(count * ratio)));
            int baseRows = Math.max(1, (int) Math.ceil(count / (double) baseColumns));

            for (int multiplier = 1; multiplier <= 12 && candidates.size() < count; multiplier++) {
                candidates.clear();
                int columns = baseColumns * multiplier;
                int rows = baseRows * multiplier;
                for (int row = 0; row < rows; row++) {
                    double y = bounds.y + bounds.height * (row + 0.5d) / rows;
                    for (int col = 0; col < columns; col++) {
                        double x = bounds.x + bounds.width * (col + 0.5d) / columns;
                        Point2D.Double point = new Point2D.Double(x, y);
                        if (isInsideBoundary(point) && isFarFromGammaMarks(point)) {
                            candidates.add(point);
                        }
                    }
                }
            }

            if (candidates.isEmpty()) {
                return List.of();
            }
            if (candidates.size() <= count) {
                return new ArrayList<>(candidates);
            }
            List<Point2D.Double> selected = new ArrayList<>();
            for (int i = 0; i < count; i++) {
                int index = count == 1
                        ? candidates.size() / 2
                        : (int) Math.round(i * (candidates.size() - 1) / (double) (count - 1));
                selected.add(candidates.get(index));
            }
            return selected;
        }

        private Rectangle2D.Double boundaryBounds() {
            double minX = Double.MAX_VALUE;
            double minY = Double.MAX_VALUE;
            double maxX = -Double.MAX_VALUE;
            double maxY = -Double.MAX_VALUE;
            for (Point2D.Double point : boundaryPoints) {
                minX = Math.min(minX, point.x);
                minY = Math.min(minY, point.y);
                maxX = Math.max(maxX, point.x);
                maxY = Math.max(maxY, point.y);
            }
            return new Rectangle2D.Double(minX, minY, maxX - minX, maxY - minY);
        }

        private boolean isInsideBoundary(Point2D.Double point) {
            boolean inside = false;
            int size = boundaryPoints.size();
            for (int i = 0, j = size - 1; i < size; j = i++) {
                Point2D.Double pi = boundaryPoints.get(i);
                Point2D.Double pj = boundaryPoints.get(j);
                boolean intersects = ((pi.y > point.y) != (pj.y > point.y))
                        && (point.x < (pj.x - pi.x) * (point.y - pi.y) / (pj.y - pi.y) + pi.x);
                if (intersects) {
                    inside = !inside;
                }
            }
            return inside;
        }

        private GammaControlPoint findGammaControlPoint(Point point) {
            if (image == null) {
                return null;
            }
            updateImageBounds();
            int radius = getGammaControlPointRadius();
            for (int i = gammaControlPoints.size() - 1; i >= 0; i--) {
                GammaControlPoint controlPoint = gammaControlPoints.get(i);
                Point center = toScreenPoint(new Point2D.Double(controlPoint.x, controlPoint.y));
                Rectangle square = new Rectangle(
                        center.x - radius,
                        center.y - radius,
                        radius * 2,
                        radius * 2
                );
                square.grow(4, 4);
                if (square.contains(point)) {
                    return controlPoint;
                }
            }
            return null;
        }

        private PprPoint findPprPoint(Point point) {
            if (image == null) {
                return null;
            }
            updateImageBounds();
            int radius = getPprPointRadius();
            for (int i = pprPoints.size() - 1; i >= 0; i--) {
                PprPoint pprPoint = pprPoints.get(i);
                Point center = toScreenPoint(new Point2D.Double(pprPoint.x, pprPoint.y));
                if (Math.abs(point.x - center.x) + Math.abs(point.y - center.y) <= radius + 4) {
                    return pprPoint;
                }
            }
            return null;
        }

        private NoisePoint findNoisePoint(Point point) {
            if (image == null) {
                return null;
            }
            updateImageBounds();
            int radius = getNoisePointRadius();
            for (int i = noisePoints.size() - 1; i >= 0; i--) {
                NoisePoint noisePoint = noisePoints.get(i);
                Point center = toScreenPoint(new Point2D.Double(noisePoint.x, noisePoint.y));
                Rectangle markerArea = new Rectangle(
                        center.x - radius - 2,
                        center.y - radius - 4,
                        radius * 2 + 42,
                        radius * 2 + 8
                );
                if (markerArea.contains(point)) {
                    return noisePoint;
                }
            }
            return null;
        }

        private int getGammaControlPointRadius() {
            return Math.max(8, Math.round(GAMMA_RADIUS * gammaControlPointScale / 100f));
        }

        private int getPprPointRadius() {
            return Math.max(8, Math.round(GAMMA_RADIUS * pprPointScale / 100f));
        }

        private int getNoisePointRadius() {
            return Math.max(6, Math.round(8f * noisePointScale / 100f));
        }

        private void renumberGammaControlPoints() {
            for (int i = 0; i < gammaControlPoints.size(); i++) {
                gammaControlPoints.get(i).number = String.valueOf(i + 1);
            }
        }

        private void renumberPprPoints() {
            for (int i = 0; i < pprPoints.size(); i++) {
                pprPoints.get(i).number = String.valueOf(i + 1);
            }
        }

        private void renumberNoisePoints() {
            for (int i = 0; i < noisePoints.size(); i++) {
                noisePoints.get(i).number = String.valueOf(i + 1);
            }
        }

        private boolean isFarFromGammaMarks(Point2D.Double point) {
            if (image == null) {
                return false;
            }
            updateImageBounds();
            Point screenPoint = toScreenPoint(point);
            for (GammaControlPoint gammaPoint : gammaControlPoints) {
                Point gammaCenter = toScreenPoint(new Point2D.Double(gammaPoint.x, gammaPoint.y));
                if (Math.abs(screenPoint.x - gammaCenter.x) < PPR_GAMMA_MIN_DISTANCE
                        && Math.abs(screenPoint.y - gammaCenter.y) < PPR_GAMMA_MIN_DISTANCE) {
                    return false;
                }
            }
            for (GammaProfile profile : gammaProfiles) {
                Point profileCenter = toScreenPoint(new Point2D.Double(profile.x, profile.y));
                if (Math.abs(screenPoint.x - profileCenter.x) < PPR_GAMMA_MIN_DISTANCE
                        && Math.abs(screenPoint.y - profileCenter.y) < PPR_GAMMA_MIN_DISTANCE) {
                    return false;
                }
            }
            return true;
        }

        private Point2D.Double toBoundaryImagePoint(Point point) {
            Point2D.Double imagePoint = toImagePoint(point);
            if (imagePoint == null) {
                return null;
            }
            if (isNearFirstBoundaryPoint(point)) {
                Point2D.Double firstPoint = boundaryPoints.get(0);
                return new Point2D.Double(firstPoint.x, firstPoint.y);
            }
            return imagePoint;
        }

        private Point snapBoundaryPreviewPoint(Point point) {
            return isNearFirstBoundaryPoint(point)
                    ? toScreenPoint(boundaryPoints.get(0))
                    : point;
        }

        private boolean isNearFirstBoundaryPoint(Point point) {
            if (boundaryPoints.size() < 3) {
                return false;
            }
            Point firstPoint = toScreenPoint(boundaryPoints.get(0));
            return point.distance(firstPoint) <= BOUNDARY_SNAP_RADIUS;
        }

        private Point2D.Double toImagePoint(Point point) {
            if (image == null) {
                return null;
            }
            updateImageBounds();
            if (!imageBounds.contains(point)) {
                return null;
            }
            double x = (point.x - imageBounds.x) * image.getWidth() / (double) imageBounds.width;
            double y = (point.y - imageBounds.y) * image.getHeight() / (double) imageBounds.height;
            return new Point2D.Double(x, y);
        }

        private Point toScreenPoint(Point2D.Double point) {
            int x = imageBounds.x + (int) Math.round(point.x * imageBounds.width / image.getWidth());
            int y = imageBounds.y + (int) Math.round(point.y * imageBounds.height / image.getHeight());
            return new Point(x, y);
        }

        private GammaProfile findGammaProfile(Point point) {
            if (image == null) {
                return null;
            }
            updateImageBounds();
            for (int i = gammaProfiles.size() - 1; i >= 0; i--) {
                GammaProfile profile = gammaProfiles.get(i);
                Point center = toScreenPoint(new Point2D.Double(profile.x, profile.y));
                int lineStartX = center.x + (int) Math.round(Math.cos(profile.angle) * GAMMA_RADIUS);
                int lineStartY = center.y + (int) Math.round(Math.sin(profile.angle) * GAMMA_RADIUS);
                int lineEndX = center.x + (int) Math.round(Math.cos(profile.angle) * (GAMMA_RADIUS + GAMMA_LINE_LENGTH));
                int lineEndY = center.y + (int) Math.round(Math.sin(profile.angle) * (GAMMA_RADIUS + GAMMA_LINE_LENGTH));
                if (point.distance(center) <= GAMMA_RADIUS + 6
                        || point.distance(lineEndX, lineEndY) <= 8
                        || distanceToSegment(point, new Point(lineStartX, lineStartY), new Point(lineEndX, lineEndY)) <= 6) {
                    return profile;
                }
            }
            return null;
        }

        private void updateGammaProfileAngle(GammaProfile profile, Point point) {
            Point center = toScreenPoint(new Point2D.Double(profile.x, profile.y));
            profile.angle = snapGammaAngle(Math.atan2(point.y - center.y, point.x - center.x));
        }

        private GammaProfile createGammaProfileFromLineEnd(Point lineEnd) {
            if (image == null) {
                return null;
            }
            updateImageBounds();
            if (!imageBounds.contains(lineEnd)) {
                return null;
            }
            GammaProfile profile = createGammaProfileFromLineEnd(lineEnd, Math.PI);
            if (profile == null) {
                profile = createGammaProfileFromLineEnd(lineEnd, 0d);
            }
            return profile;
        }

        private GammaProfile createGammaProfileFromLineEnd(Point lineEnd, double angle) {
            int centerX = lineEnd.x - (int) Math.round(Math.cos(angle) * (GAMMA_RADIUS + GAMMA_LINE_LENGTH));
            int centerY = lineEnd.y - (int) Math.round(Math.sin(angle) * (GAMMA_RADIUS + GAMMA_LINE_LENGTH));
            Point2D.Double centerPoint = toImagePoint(new Point(centerX, centerY));
            return centerPoint == null ? null : new GammaProfile(centerPoint.x, centerPoint.y, angle);
        }

        private boolean editGammaProfileNumber(GammaProfile profile) {
            while (true) {
                JTextField field = new JTextField(profile.number, 8);
                field.selectAll();
                JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
                panel.add(new JLabel("Номер профиля:"));
                panel.add(field);
                panel.addHierarchyListener(e -> {
                    if (panel.isShowing()) {
                        SwingUtilities.invokeLater(() -> {
                            field.requestFocusInWindow();
                            field.selectAll();
                        });
                    }
                });
                int result = JOptionPane.showOptionDialog(
                        this,
                        panel,
                        "Номер профиля",
                        JOptionPane.YES_NO_OPTION,
                        JOptionPane.QUESTION_MESSAGE,
                        null,
                        new Object[]{"Да", "Отмена"},
                        "Да"
                );
                if (result != JOptionPane.YES_OPTION) {
                    return false;
                }
                String trimmed = field.getText().trim();
                if (trimmed.matches("\\d{1,3}")) {
                    profile.number = trimmed;
                    fireGammaProfilesChanged();
                    repaint();
                    return true;
                }
                JOptionPane.showOptionDialog(this,
                        "Введите число от 1 до 999.",
                        "Номер профиля",
                        JOptionPane.DEFAULT_OPTION,
                        JOptionPane.WARNING_MESSAGE,
                        null,
                        new Object[]{"Да"},
                        "Да");
            }
        }

        private double snapGammaAngle(double angle) {
            return Math.round(angle / GAMMA_ANGLE_STEP) * GAMMA_ANGLE_STEP;
        }

        private double distanceToSegment(Point point, Point start, Point end) {
            double dx = end.x - start.x;
            double dy = end.y - start.y;
            if (dx == 0d && dy == 0d) {
                return point.distance(start);
            }
            double t = ((point.x - start.x) * dx + (point.y - start.y) * dy) / (dx * dx + dy * dy);
            t = Math.max(0d, Math.min(1d, t));
            double projectionX = start.x + t * dx;
            double projectionY = start.y + t * dy;
            return point.distance(projectionX, projectionY);
        }

        @Override
        public Dimension getPreferredScrollableViewportSize() {
            return getPreferredSize();
        }

        @Override
        public int getScrollableUnitIncrement(Rectangle visibleRect, int orientation, int direction) {
            return 32;
        }

        @Override
        public int getScrollableBlockIncrement(Rectangle visibleRect, int orientation, int direction) {
            return 128;
        }

        @Override
        public boolean getScrollableTracksViewportWidth() {
            return true;
        }

        @Override
        public boolean getScrollableTracksViewportHeight() {
            return true;
        }
    }

    public AreaProtocolPanel() {
        super(new BorderLayout());
        setBorder(BorderFactory.createEmptyBorder(4, 4, 4, 4));

        JPanel formPanel = new JPanel(new GridBagLayout());
        formPanel.setBorder(new TitledBorder("Реквизиты протокола и заказчика"));

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4, 4, 4, 4);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 0.0;
        gbc.gridy = 0;

        styleInput(projectNameField);
        addLabeledField(formPanel, gbc, "Название проекта:", projectNameField);

        protocolDatePicker = createDatePicker(LocalDate.now().plusDays(1));
        addLabeledField(formPanel, gbc, "Дата протокола:", protocolDatePicker);

        customerNameContactsField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Наименование и контактные данные заказчика:",
                customerNameContactsField);
        installPlaceholder(customerNameContactsField, CUSTOMER_PLACEHOLDER);

        customerLegalAddressField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Юридический адрес заказчика:",
                customerLegalAddressField);

        customerActualAddressField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Фактический адрес заказчика:",
                customerActualAddressField);

        objectNameArea = new JTextArea(3, 40);
        objectNameArea.setLineWrap(true);
        objectNameArea.setWrapStyleWord(true);
        JScrollPane objectNameScroll = new JScrollPane(
                objectNameArea,
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                JScrollPane.HORIZONTAL_SCROLLBAR_NEVER
        );
        addLabeledField(formPanel, gbc,
                "Наименование объекта измерений:",
                objectNameScroll);

        objectAddressField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Адрес предприятия (объекта):",
                objectAddressField);

        contractNumberField = new JTextField(10);
        contractDatePicker = createDatePicker(null);
        JPanel contractPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        contractPanel.add(new JLabel("№"));
        contractPanel.add(contractNumberField);
        contractPanel.add(new JLabel("от"));
        contractPanel.add(contractDatePicker);
        addLabeledField(formPanel, gbc, "Договор:", contractPanel);

        applicationNumberField = new JTextField(10);
        applicationDatePicker = createDatePicker(null);
        JPanel applicationPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        applicationPanel.add(new JLabel("№"));
        applicationPanel.add(applicationNumberField);
        applicationPanel.add(new JLabel("от"));
        applicationPanel.add(applicationDatePicker);
        addLabeledField(formPanel, gbc, "Заявка:", applicationPanel);

        representativeField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Представитель заказчика (ФИО, должность):",
                representativeField);

        JPanel areaPanel = createAreaPanel();

        JPanel bottomPanel = new JPanel(new BorderLayout());
        bottomPanel.setBorder(new TitledBorder("Даты проведения измерений и температуры воздуха"));

        measurementRowsPanel = new JPanel();
        measurementRowsPanel.setLayout(new BoxLayout(measurementRowsPanel, BoxLayout.Y_AXIS));
        JScrollPane measurementScroll = new JScrollPane(
                measurementRowsPanel,
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                JScrollPane.HORIZONTAL_SCROLLBAR_NEVER
        );

        JPanel buttonsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        JButton addRowButton = new JButton("Добавить дату измерений");
        JButton removeRowButton = new JButton("Удалить последнюю строку");
        addRowButton.addActionListener(e -> addMeasurementRow(null));
        removeRowButton.addActionListener(e -> removeLastMeasurementRow());
        buttonsPanel.add(addRowButton);
        buttonsPanel.add(removeRowButton);

        addMeasurementRow(null);

        bottomPanel.add(buttonsPanel, BorderLayout.NORTH);
        bottomPanel.add(measurementScroll, BorderLayout.CENTER);

        JPanel centerPanel = new JPanel(new BorderLayout());
        centerPanel.add(areaPanel, BorderLayout.NORTH);
        centerPanel.add(bottomPanel, BorderLayout.CENTER);

        JPanel radiationTitleTab = new JPanel(new BorderLayout());
        radiationTitleTab.add(formPanel, BorderLayout.NORTH);
        radiationTitleTab.add(centerPanel, BorderLayout.CENTER);
        radiationTitleTab.add(createImportPanel(), BorderLayout.SOUTH);

        JTabbedPane tabs = new JTabbedPane(JTabbedPane.TOP);
        tabs.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
        tabs.addTab("Титульная страница", radiationTitleTab);
        tabs.addTab("Эскиз", createSketchPanel());
        tabs.addTab("Значения", createValuesPanel());

        add(tabs, BorderLayout.CENTER);
    }

    public AreaProtocolData getProtocolData() {
        return new AreaProtocolData(
                getDateText(protocolDatePicker),
                getFieldText(customerNameContactsField, CUSTOMER_PLACEHOLDER),
                customerLegalAddressField.getText().trim(),
                customerActualAddressField.getText().trim(),
                objectNameArea.getText().trim(),
                objectAddressField.getText().trim(),
                contractNumberField.getText().trim(),
                getDateText(contractDatePicker),
                applicationNumberField.getText().trim(),
                getDateText(applicationDatePicker),
                representativeField.getText().trim(),
                collectMeasurementRows(),
                collectMedValues(),
                gammaMinValueField.getText().trim(),
                gammaMaxValueField.getText().trim(),
                gammaAverageValueField.getText().trim(),
                pprMinValueField.getText().trim(),
                pprMaxValueField.getText().trim(),
                noiseEquivalentMinValueField.getText().trim(),
                noiseEquivalentMaxValueField.getText().trim(),
                noiseMaxLevelMinValueField.getText().trim(),
                noiseMaxLevelMaxValueField.getText().trim(),
                String.valueOf(noiseMethodComboBox.getSelectedItem()),
                areaField.getText().trim(),
                medCheckBox.isSelected(),
                pprCheckBox.isSelected(),
                sketchPreviewPanel.renderSketchImage(false),
                sketchPreviewPanel.renderNoiseSketchImage(),
                sketchPreviewPanel.getMaxGammaProfileNumber(),
                sketchPreviewPanel.getGammaControlPointCount(),
                sketchPreviewPanel.getPprPointCount(),
                sketchPreviewPanel.getNoisePointCount(),
                imageFile
        );
    }

    private JPanel createAreaPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(new TitledBorder("Данные участка"));

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4, 4, 4, 4);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.gridy = 0;

        areaField.setColumns(18);
        styleInput(areaField);

        gbc.gridx = 0;
        gbc.weightx = 0.0;
        panel.add(new JLabel("Площадь участка:"), gbc);
        gbc.gridx = 1;
        gbc.weightx = 1.0;
        panel.add(areaField, gbc);
        gbc.gridx = 2;
        gbc.weightx = 0.0;
        panel.add(new JLabel("м²"), gbc);

        JPanel checks = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 0));
        checks.add(medCheckBox);
        checks.add(pprCheckBox);

        gbc.gridx = 0;
        gbc.gridy++;
        panel.add(new JLabel("Измеряем:"), gbc);
        gbc.gridx = 1;
        gbc.gridwidth = 2;
        panel.add(checks, gbc);
        gbc.gridwidth = 1;

        return panel;
    }

    private JPanel createSketchPanel() {
        JPanel panel = new JPanel(new BorderLayout(8, 8));
        panel.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));

        JButton imageButton = new JButton("Добавить эскиз");
        imageButton.addActionListener(e -> chooseImage());

        JToggleButton boundaryButton = new JToggleButton("Граница участка");
        JToggleButton gammaProfileButton = new JToggleButton("Профили гамма");
        JToggleButton gammaControlPointButton = new JToggleButton("КТ гамма вручную");
        JToggleButton pprPointButton = new JToggleButton("ППР вручную");
        JToggleButton noisePointButton = new JToggleButton("Шум вручную");
        JToggleButton radiationLayerButton = new JToggleButton("Радиация", true);
        JToggleButton noiseLayerButton = new JToggleButton("Шум", true);
        radiationLayerButton.addActionListener(e ->
                sketchPreviewPanel.setRadiationLayerVisible(radiationLayerButton.isSelected()));
        noiseLayerButton.addActionListener(e ->
                sketchPreviewPanel.setNoiseLayerVisible(noiseLayerButton.isSelected()));
        sketchPreviewPanel.setBoundaryModeOffAction(() -> {
            boundaryButton.setSelected(false);
            sketchPreviewPanel.setBoundaryMode(false);
        });
        sketchPreviewPanel.setGammaProfilesChangedAction(this::syncMedValueRowsToGammaProfiles);
        boundaryButton.addActionListener(e -> {
            if (boundaryButton.isSelected()) {
                gammaProfileButton.setSelected(false);
                gammaControlPointButton.setSelected(false);
                pprPointButton.setSelected(false);
                noisePointButton.setSelected(false);
                sketchPreviewPanel.setGammaProfileMode(false);
                sketchPreviewPanel.setGammaControlPointMode(false);
                sketchPreviewPanel.setPprPointMode(false);
                sketchPreviewPanel.setNoisePointMode(false);
            }
            sketchPreviewPanel.setBoundaryMode(boundaryButton.isSelected());
        });
        gammaProfileButton.addActionListener(e -> {
            if (gammaProfileButton.isSelected()) {
                radiationLayerButton.setSelected(true);
                sketchPreviewPanel.setRadiationLayerVisible(true);
                boundaryButton.setSelected(false);
                gammaControlPointButton.setSelected(false);
                pprPointButton.setSelected(false);
                noisePointButton.setSelected(false);
                sketchPreviewPanel.setBoundaryMode(false);
                sketchPreviewPanel.setGammaControlPointMode(false);
                sketchPreviewPanel.setPprPointMode(false);
                sketchPreviewPanel.setNoisePointMode(false);
            }
            sketchPreviewPanel.setGammaProfileMode(gammaProfileButton.isSelected());
        });
        gammaControlPointButton.addActionListener(e -> {
            if (gammaControlPointButton.isSelected()) {
                radiationLayerButton.setSelected(true);
                sketchPreviewPanel.setRadiationLayerVisible(true);
                boundaryButton.setSelected(false);
                gammaProfileButton.setSelected(false);
                pprPointButton.setSelected(false);
                noisePointButton.setSelected(false);
                sketchPreviewPanel.setBoundaryMode(false);
                sketchPreviewPanel.setGammaProfileMode(false);
                sketchPreviewPanel.setPprPointMode(false);
                sketchPreviewPanel.setNoisePointMode(false);
            }
            sketchPreviewPanel.setGammaControlPointMode(gammaControlPointButton.isSelected());
        });
        pprPointButton.addActionListener(e -> {
            if (pprPointButton.isSelected()) {
                radiationLayerButton.setSelected(true);
                sketchPreviewPanel.setRadiationLayerVisible(true);
                boundaryButton.setSelected(false);
                gammaProfileButton.setSelected(false);
                gammaControlPointButton.setSelected(false);
                noisePointButton.setSelected(false);
                sketchPreviewPanel.setBoundaryMode(false);
                sketchPreviewPanel.setGammaProfileMode(false);
                sketchPreviewPanel.setGammaControlPointMode(false);
                sketchPreviewPanel.setNoisePointMode(false);
            }
            sketchPreviewPanel.setPprPointMode(pprPointButton.isSelected());
        });
        noisePointButton.addActionListener(e -> {
            if (noisePointButton.isSelected()) {
                noiseLayerButton.setSelected(true);
                sketchPreviewPanel.setNoiseLayerVisible(true);
                boundaryButton.setSelected(false);
                gammaProfileButton.setSelected(false);
                gammaControlPointButton.setSelected(false);
                pprPointButton.setSelected(false);
                sketchPreviewPanel.setBoundaryMode(false);
                sketchPreviewPanel.setGammaProfileMode(false);
                sketchPreviewPanel.setGammaControlPointMode(false);
                sketchPreviewPanel.setPprPointMode(false);
            }
            sketchPreviewPanel.setNoisePointMode(noisePointButton.isSelected());
        });

        JComboBox<String> colorBox = new JComboBox<>(new String[]{"Черный", "Белый"});
        colorBox.addActionListener(e -> sketchPreviewPanel.setBoundaryColor(
                colorBox.getSelectedIndex() == 1 ? Color.WHITE : Color.BLACK
        ));

        JButton gammaControlPointsButton = new JButton("Вставить КТ гамма");
        gammaControlPointsButton.addActionListener(e -> insertGammaControlPoints());

        JButton pprPointsButton = new JButton("Вставить ППР");
        pprPointsButton.addActionListener(e -> insertPprPoints());

        gammaControlPointSizeSlider = new JSlider(60, 160, 100);
        gammaControlPointSizeSlider.setPreferredSize(new Dimension(120, 28));
        gammaControlPointSizeSlider.setToolTipText("Размер квадратиков КТ гамма");
        gammaControlPointSizeSlider.addChangeListener(e ->
                sketchPreviewPanel.setGammaControlPointScale(gammaControlPointSizeSlider.getValue()));

        pprPointSizeSlider = new JSlider(60, 160, 100);
        pprPointSizeSlider.setPreferredSize(new Dimension(120, 28));
        pprPointSizeSlider.setToolTipText("Размер ромбиков ППР");
        pprPointSizeSlider.addChangeListener(e ->
                sketchPreviewPanel.setPprPointScale(pprPointSizeSlider.getValue()));

        noisePointSizeSlider = new JSlider(60, 160, 100);
        noisePointSizeSlider.setPreferredSize(new Dimension(120, 28));
        noisePointSizeSlider.setToolTipText("Размер точек шума");
        noisePointSizeSlider.addChangeListener(e ->
                sketchPreviewPanel.setNoisePointScale(noisePointSizeSlider.getValue()));

        JButton helpButton = new JButton("?");
        helpButton.setToolTipText("Подсказки");
        helpButton.addActionListener(e -> showSketchHelp());

        JPanel sketchBlock = createSketchControlGroup("Эскиз");
        addSketchControl(sketchBlock, imageButton);
        imageLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
        sketchBlock.add(Box.createVerticalStrut(6));
        sketchBlock.add(imageLabel);

        JPanel layerBlock = createSketchControlGroup("Слои");
        addSketchControl(layerBlock, radiationLayerButton);
        addSketchControl(layerBlock, noiseLayerButton);

        JPanel boundaryBlock = createSketchControlGroup("Граница участка");
        addSketchControl(boundaryBlock, boundaryButton);
        JPanel colorPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        colorPanel.setOpaque(false);
        colorPanel.setAlignmentX(Component.LEFT_ALIGNMENT);
        colorPanel.add(new JLabel("Цвет:"));
        colorPanel.add(colorBox);
        boundaryBlock.add(Box.createVerticalStrut(6));
        boundaryBlock.add(colorPanel);

        JPanel profileBlock = createSketchControlGroup("Профили");
        addSketchControl(profileBlock, gammaProfileButton);

        JPanel manualBlock = createSketchControlGroup("Вручную");
        addSketchControl(manualBlock, gammaControlPointButton);
        addSketchControl(manualBlock, pprPointButton);
        addSketchControl(manualBlock, noisePointButton);

        JPanel autoBlock = createSketchControlGroup("Авто-вставка");
        addSketchControl(autoBlock, gammaControlPointsButton);
        addSketchControl(autoBlock, pprPointsButton);

        JPanel sizeBlock = createSketchControlGroup("Размеры");
        addSketchSlider(sizeBlock, "КТ гамма:", gammaControlPointSizeSlider);
        addSketchSlider(sizeBlock, "ППР:", pprPointSizeSlider);
        addSketchSlider(sizeBlock, "Шум:", noisePointSizeSlider);

        JPanel helpBlock = createSketchControlGroup("Помощь");
        addSketchControl(helpBlock, helpButton);

        lockSketchControlGroupHeight(sketchBlock);
        lockSketchControlGroupHeight(layerBlock);
        lockSketchControlGroupHeight(boundaryBlock);
        lockSketchControlGroupHeight(profileBlock);
        lockSketchControlGroupHeight(manualBlock);
        lockSketchControlGroupHeight(autoBlock);
        lockSketchControlGroupHeight(sizeBlock);
        lockSketchControlGroupHeight(helpBlock);

        JPanel controlsPanel = new JPanel();
        controlsPanel.setLayout(new BoxLayout(controlsPanel, BoxLayout.Y_AXIS));
        controlsPanel.setBorder(BorderFactory.createEmptyBorder(0, 8, 0, 0));
        controlsPanel.add(sketchBlock);
        controlsPanel.add(Box.createVerticalStrut(8));
        controlsPanel.add(layerBlock);
        controlsPanel.add(Box.createVerticalStrut(8));
        controlsPanel.add(boundaryBlock);
        controlsPanel.add(Box.createVerticalStrut(8));
        controlsPanel.add(profileBlock);
        controlsPanel.add(Box.createVerticalStrut(8));
        controlsPanel.add(manualBlock);
        controlsPanel.add(Box.createVerticalStrut(8));
        controlsPanel.add(autoBlock);
        controlsPanel.add(Box.createVerticalStrut(8));
        controlsPanel.add(sizeBlock);
        controlsPanel.add(Box.createVerticalStrut(8));
        controlsPanel.add(helpBlock);
        controlsPanel.add(Box.createVerticalGlue());

        JScrollPane controlsScroll = new JScrollPane(
                controlsPanel,
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                JScrollPane.HORIZONTAL_SCROLLBAR_NEVER
        );
        controlsScroll.setBorder(BorderFactory.createEmptyBorder());
        controlsScroll.setPreferredSize(new Dimension(260, 0));

        panel.add(new JScrollPane(sketchPreviewPanel), BorderLayout.CENTER);
        panel.add(controlsScroll, BorderLayout.EAST);
        return panel;
    }

    private JPanel createSketchControlGroup(String title) {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBorder(BorderFactory.createTitledBorder(title));
        panel.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.setMaximumSize(new Dimension(Integer.MAX_VALUE, Integer.MAX_VALUE));
        return panel;
    }

    private void lockSketchControlGroupHeight(JPanel panel) {
        panel.setMaximumSize(new Dimension(Integer.MAX_VALUE, panel.getPreferredSize().height));
    }

    private void addSketchControl(JPanel panel, JComponent component) {
        component.setAlignmentX(Component.LEFT_ALIGNMENT);
        Dimension preferred = component.getPreferredSize();
        component.setMaximumSize(new Dimension(Integer.MAX_VALUE, preferred.height));
        panel.add(component);
        panel.add(Box.createVerticalStrut(6));
    }

    private void addSketchSlider(JPanel panel, String label, JSlider slider) {
        JLabel title = new JLabel(label);
        title.setAlignmentX(Component.LEFT_ALIGNMENT);
        slider.setAlignmentX(Component.LEFT_ALIGNMENT);
        slider.setMaximumSize(new Dimension(Integer.MAX_VALUE, slider.getPreferredSize().height));
        panel.add(title);
        panel.add(slider);
        panel.add(Box.createVerticalStrut(6));
    }

    private void showSketchHelp() {
        JOptionPane.showMessageDialog(this,
                "Граница участка: включите кнопку и кликайте по эскизу.\n"
                        + "Слои показывают или скрывают радиацию и шум на эскизе.\n"
                        + "Профили гамма: клик ставит конец палочки, затем вводится номер профиля.\n"
                        + "Двойной клик по кружку меняет номер.\n"
                        + "Перетаскивание кружка или палочки поворачивает профиль.\n"
                        + "Стрелки двигают выбранный профиль гамма, КТ гамма, ППР или шум на 1 пиксель.\n"
                        + "Вставить КТ гамма расставляет квадратные контрольные точки внутри границы участка.\n"
                        + "КТ гамма вручную: включите режим и кликайте внутри границы участка.\n"
                        + "Вставить ППР расставляет ромбики ППР внутри границы участка.\n"
                        + "ППР вручную: включите режим и кликайте внутри границы участка.\n"
                        + "Шум вручную: включите режим и кликайте внутри границы участка.\n"
                        + "Размер КТ меняет размер квадратных контрольных точек.\n"
                        + "Размер ППР меняет размер ромбиков ППР.\n"
                        + "Размер шума меняет размер черных точек шума.\n"
                        + "В ручных режимах КТ гамма, ППР и шум можно перетаскивать левой кнопкой мыши.\n"
                        + "Ctrl+Z отменяет последнее действие на эскизе.\n"
                        + "Delete удаляет наведенную/выбранную КТ гамма, ППР, шум или выбранный профиль гамма.",
                "Подсказки",
                JOptionPane.INFORMATION_MESSAGE);
    }

    private JPanel createValuesPanel() {
        JPanel panel = new JPanel(new BorderLayout(8, 8));
        panel.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
        JComboBox<String> selector = new JComboBox<>(new String[]{"МЭД", "ППР", "Шум"});
        JPanel cards = new JPanel(new CardLayout());
        cards.add(new JScrollPane(createMedValuesSection()), "МЭД");
        cards.add(createValuesSection("ППР", new String[]{""},
                new JTextField[]{pprMinValueField},
                new JTextField[]{pprMaxValueField}), "ППР");
        cards.add(createNoiseValuesSection(), "Шум");
        selector.addActionListener(e ->
                ((CardLayout) cards.getLayout()).show(cards, (String) selector.getSelectedItem()));
        panel.add(selector, BorderLayout.NORTH);
        panel.add(cards, BorderLayout.CENTER);
        return panel;
    }

    private JPanel createMedValuesSection() {
        JPanel panel = new JPanel(new BorderLayout(6, 6));
        panel.setBorder(BorderFactory.createTitledBorder("МЭД"));
        panel.setAlignmentX(Component.LEFT_ALIGNMENT);

        JPanel topPanel = new JPanel();
        topPanel.setLayout(new BoxLayout(topPanel, BoxLayout.Y_AXIS));
        topPanel.add(createMedSummaryPanel());
        applyMedCountToAllCheckBox.addActionListener(e -> {
            if (applyMedCountToAllCheckBox.isSelected()) {
                applyFirstMedCountToAllRows();
            }
        });
        topPanel.add(applyMedCountToAllCheckBox);
        panel.add(topPanel, BorderLayout.NORTH);
        panel.add(medValuesRowsPanel, BorderLayout.CENTER);
        syncMedValueRowsToGammaProfiles();
        return panel;
    }

    private JPanel createMedSummaryPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
        panel.setAlignmentX(Component.LEFT_ALIGNMENT);
        styleCompactMedField(gammaMinValueField);
        styleCompactMedField(gammaMaxValueField);
        styleCompactMedField(gammaAverageValueField);
        panel.add(new JLabel("Минимум:"));
        panel.add(gammaMinValueField);
        panel.add(new JLabel("Максимум:"));
        panel.add(gammaMaxValueField);
        panel.add(new JLabel("Среднее:"));
        panel.add(gammaAverageValueField);
        return panel;
    }

    private JPanel createNoiseValuesSection() {
        JPanel panel = createValuesSection("Шум", new String[]{
                "Эквивалентные уровни звука:",
                "Максимальные уровни звука:"
        }, new JTextField[]{
                noiseEquivalentMinValueField,
                noiseMaxLevelMinValueField
        }, new JTextField[]{
                noiseEquivalentMaxValueField,
                noiseMaxLevelMaxValueField
        });
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridy = 2;
        gbc.insets = new Insets(4, 6, 4, 6);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.gridx = 0;
        gbc.fill = GridBagConstraints.NONE;
        gbc.weightx = 0.0;
        panel.add(new JLabel("Методика:"), gbc);
        gbc.gridx = 1;
        gbc.gridwidth = 4;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        panel.add(noiseMethodComboBox, gbc);
        return panel;
    }

    private void syncMedValueRowsToGammaProfiles() {
        int profileCount = sketchPreviewPanel.getMaxGammaProfileNumber();
        if (profileCount == medValueRows.size() && medValuesRowsPanel.getComponentCount() > 0) {
            return;
        }
        List<MedValueSnapshot> oldValues = collectMedValueSnapshots();
        medValueRows.clear();
        medValuesRowsPanel.removeAll();
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4, 6, 4, 6);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.gridy = 0;

        int columnCount = profileCount > 20 ? 3 : profileCount > 10 ? 2 : 1;
        int rowsPerColumn = Math.max(1, (int) Math.ceil(profileCount / (double) columnCount));
        for (int columnBlock = 0; columnBlock < columnCount; columnBlock++) {
            int baseColumn = columnBlock * 4;
            addMedHeader(gbc, baseColumn, "Профиль");
            addMedHeader(gbc, baseColumn + 1, "Расстояние, м");
            addMedHeader(gbc, baseColumn + 2, "Кол-во");
        }

        for (int profileNumber = 1; profileNumber <= profileCount; profileNumber++) {
            MedValueRow row = new MedValueRow();
            row.profileNumber = profileNumber;
            row.distanceField = createCompactValueField();
            row.countField = createCompactValueField();
            applySavedMedValue(row, oldValues);
            addMedRowListeners(row);

            int columnBlock = (profileNumber - 1) / rowsPerColumn;
            int baseColumn = columnBlock * 4;
            gbc.gridy = (profileNumber - 1) % rowsPerColumn + 1;
            gbc.gridx = baseColumn;
            medValuesRowsPanel.add(new JLabel("Профиль " + profileNumber), gbc);
            gbc.gridx = baseColumn + 1;
            medValuesRowsPanel.add(row.distanceField, gbc);
            gbc.gridx = baseColumn + 2;
            medValuesRowsPanel.add(row.countField, gbc);
            medValueRows.add(row);
        }
        if (profileCount == 0) {
            gbc.gridy = 1;
            gbc.gridx = 0;
            gbc.gridwidth = 3;
            medValuesRowsPanel.add(new JLabel("Добавьте профили гамма на эскизе."), gbc);
        }
        updateMedFocusTraversalPolicy();
        medValuesRowsPanel.revalidate();
        medValuesRowsPanel.repaint();
    }

    private void updateMedFocusTraversalPolicy() {
        List<Component> focusOrder = new ArrayList<>();
        for (MedValueRow row : medValueRows) {
            focusOrder.add(row.distanceField);
            focusOrder.add(row.countField);
        }
        medValuesRowsPanel.setFocusTraversalPolicyProvider(!focusOrder.isEmpty());
        if (focusOrder.isEmpty()) {
            return;
        }
        medValuesRowsPanel.setFocusTraversalPolicy(new FocusTraversalPolicy() {
            @Override
            public Component getComponentAfter(Container container, Component component) {
                int index = focusOrder.indexOf(component);
                return focusOrder.get(index < 0 || index + 1 >= focusOrder.size() ? 0 : index + 1);
            }

            @Override
            public Component getComponentBefore(Container container, Component component) {
                int index = focusOrder.indexOf(component);
                return focusOrder.get(index <= 0 ? focusOrder.size() - 1 : index - 1);
            }

            @Override
            public Component getFirstComponent(Container container) {
                return focusOrder.get(0);
            }

            @Override
            public Component getLastComponent(Container container) {
                return focusOrder.get(focusOrder.size() - 1);
            }

            @Override
            public Component getDefaultComponent(Container container) {
                return getFirstComponent(container);
            }
        });
    }

    private void addMedHeader(GridBagConstraints gbc, int column, String text) {
        gbc.gridx = column;
        medValuesRowsPanel.add(new JLabel(text), gbc);
    }

    private JTextField createCompactValueField() {
        JTextField field = new JTextField(6);
        styleCompactMedField(field);
        return field;
    }

    private void styleCompactMedField(JTextField field) {
        field.setPreferredSize(new Dimension(72, 26));
        field.setMinimumSize(new Dimension(64, 26));
        field.setMargin(new Insets(1, 4, 1, 4));
    }

    private void applySavedMedValue(MedValueRow row, List<MedValueSnapshot> values) {
        for (MedValueSnapshot value : values) {
            if (value.profileNumber == row.profileNumber) {
                setFieldText(row.distanceField, value.distance);
                String count = value.count;
                if ((count == null || count.isBlank()) && value.distance != null && !value.distance.isBlank()) {
                    count = calculateMedMeasurementCount(value.distance);
                }
                setFieldText(row.countField, count);
                return;
            }
        }
    }

    private void addMedRowListeners(MedValueRow row) {
        addTextChangeListener(row.distanceField, () -> updateMedCountFromDistance(row));
        addTextChangeListener(row.countField, () -> {
            if (!updatingMedValues && applyMedCountToAllCheckBox.isSelected()) {
                applyMedCountToAllRows(row);
            }
        });
    }

    private void updateMedCountFromDistance(MedValueRow row) {
        if (updatingMedValues) {
            return;
        }
        String count = calculateMedMeasurementCount(row.distanceField.getText());
        updatingMedValues = true;
        row.countField.setText(count);
        updatingMedValues = false;
        if (applyMedCountToAllCheckBox.isSelected()) {
            applyMedCountToAllRows(row);
        }
    }

    private String calculateMedMeasurementCount(String distanceText) {
        double distance = parsePositiveDouble(distanceText);
        if (distance <= 0d) {
            return "";
        }
        double metersPerMeasurement = 5000d / 3600d * 8d;
        return String.valueOf(Math.max(1, (int) Math.ceil(distance / metersPerMeasurement)));
    }

    private double parsePositiveDouble(String text) {
        String normalized = text == null ? "" : text.trim().replace(" ", "").replace(",", ".");
        if (normalized.isBlank()) {
            return 0d;
        }
        try {
            return Math.max(0d, Double.parseDouble(normalized));
        } catch (NumberFormatException ex) {
            return 0d;
        }
    }

    private void applyFirstMedCountToAllRows() {
        if (!medValueRows.isEmpty()) {
            applyMedCountToAllRows(medValueRows.get(0));
        }
    }

    private void applyMedCountToAllRows(MedValueRow source) {
        updatingMedValues = true;
        for (MedValueRow row : medValueRows) {
            if (row == source) {
                continue;
            }
            row.distanceField.setText(source.distanceField.getText());
            row.countField.setText(source.countField.getText());
        }
        updatingMedValues = false;
    }

    private void addTextChangeListener(JTextField field, Runnable action) {
        field.getDocument().addDocumentListener(new DocumentListener() {
            @Override
            public void insertUpdate(DocumentEvent e) {
                action.run();
            }

            @Override
            public void removeUpdate(DocumentEvent e) {
                action.run();
            }

            @Override
            public void changedUpdate(DocumentEvent e) {
                action.run();
            }
        });
    }

    private List<MedValueSnapshot> collectMedValueSnapshots() {
        List<MedValueSnapshot> values = new ArrayList<>();
        for (MedValueRow row : medValueRows) {
            MedValueSnapshot value = new MedValueSnapshot();
            value.profileNumber = row.profileNumber;
            value.distance = row.distanceField.getText().trim();
            value.count = row.countField.getText().trim();
            values.add(value);
        }
        return values;
    }

    private void applyMedValueSnapshots(List<MedValueSnapshot> values, String fallbackMin,
                                        String fallbackMax, String fallbackAverage) {
        setFieldText(gammaMinValueField, fallbackMin);
        setFieldText(gammaMaxValueField, fallbackMax);
        setFieldText(gammaAverageValueField, fallbackAverage);
        syncMedValueRowsToGammaProfiles();
        if (values != null && !values.isEmpty()) {
            updatingMedValues = true;
            for (MedValueRow row : medValueRows) {
                applySavedMedValue(row, values);
            }
            updatingMedValues = false;
        }
    }

    private JPanel createValuesSection(String title, String[] rowLabels,
                                       JTextField[] minFields, JTextField[] maxFields) {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(title));
        panel.setAlignmentX(Component.LEFT_ALIGNMENT);

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridy = 0;
        gbc.insets = new Insets(4, 6, 4, 6);
        gbc.anchor = GridBagConstraints.WEST;

        for (int i = 0; i < rowLabels.length; i++) {
            int column = 0;
            String rowLabel = rowLabels[i] == null ? "" : rowLabels[i];
            if (!rowLabel.isBlank()) {
                gbc.gridx = column++;
                gbc.fill = GridBagConstraints.NONE;
                gbc.weightx = 0.0;
                panel.add(new JLabel(rowLabel), gbc);
            }
            gbc.gridx = column++;
            panel.add(new JLabel("Минимум:"), gbc);
            gbc.gridx = column++;
            panel.add(minFields[i], gbc);
            gbc.gridx = column++;
            panel.add(new JLabel("Максимум:"), gbc);
            gbc.gridx = column;
            panel.add(maxFields[i], gbc);
            gbc.gridy++;
        }

        return panel;
    }

    private void insertGammaControlPoints() {
        Integer count = requestGammaControlPointCount();
        if (count == null) {
            return;
        }
        sketchPreviewPanel.insertGammaControlPoints(count);
    }

    private void insertPprPoints() {
        Integer count = requestPprPointCount();
        if (count == null) {
            return;
        }
        sketchPreviewPanel.insertPprPoints(count);
    }

    private Integer requestGammaControlPointCount() {
        Integer calculatedCount = calculateGammaControlPointCount();
        JTextField field = new JTextField(calculatedCount == null ? "" : String.valueOf(calculatedCount), 8);
        field.selectAll();

        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
        panel.add(new JLabel("Количество КТ гамма:"));
        panel.add(field);

        int result = JOptionPane.showOptionDialog(
                this,
                panel,
                "КТ гамма",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                new Object[]{"Да", "Отмена"},
                "Да"
        );
        if (result != JOptionPane.YES_OPTION) {
            return null;
        }
        try {
            int count = Integer.parseInt(field.getText().trim());
            if (count < 5) {
                JOptionPane.showOptionDialog(this,
                        "Количество КТ гамма не может быть меньше 5.",
                        "КТ гамма",
                        JOptionPane.DEFAULT_OPTION,
                        JOptionPane.WARNING_MESSAGE,
                        null,
                        new Object[]{"Да"},
                        "Да");
                return null;
            }
            return count;
        } catch (NumberFormatException ex) {
            JOptionPane.showOptionDialog(this,
                    "Введите целое число.",
                    "КТ гамма",
                    JOptionPane.DEFAULT_OPTION,
                    JOptionPane.WARNING_MESSAGE,
                    null,
                    new Object[]{"Да"},
                    "Да");
            return null;
        }
    }

    private Integer requestPprPointCount() {
        Integer minimumCount = calculatePprPointCount();
        JTextField field = new JTextField(minimumCount == null ? "" : String.valueOf(minimumCount), 8);
        field.selectAll();

        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
        panel.add(new JLabel("Количество ППР:"));
        panel.add(field);

        int result = JOptionPane.showOptionDialog(
                this,
                panel,
                "ППР",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                new Object[]{"Да", "Отмена"},
                "Да"
        );
        if (result != JOptionPane.YES_OPTION) {
            return null;
        }
        try {
            int count = Integer.parseInt(field.getText().trim());
            if (count <= 0) {
                JOptionPane.showOptionDialog(this,
                        "Введите количество ППР больше 0.",
                        "ППР",
                        JOptionPane.DEFAULT_OPTION,
                        JOptionPane.WARNING_MESSAGE,
                        null,
                        new Object[]{"Да"},
                        "Да");
                return null;
            }
            if (minimumCount != null && count < minimumCount) {
                JOptionPane.showOptionDialog(this,
                        "По указанной площади нужно минимум " + minimumCount + " точек ППР.",
                        "ППР",
                        JOptionPane.DEFAULT_OPTION,
                        JOptionPane.WARNING_MESSAGE,
                        null,
                        new Object[]{"Да"},
                        "Да");
                return null;
            }
            return count;
        } catch (NumberFormatException ex) {
            JOptionPane.showOptionDialog(this,
                    "Введите целое число.",
                    "ППР",
                    JOptionPane.DEFAULT_OPTION,
                    JOptionPane.WARNING_MESSAGE,
                    null,
                    new Object[]{"Да"},
                    "Да");
            return null;
        }
    }

    private Integer calculateGammaControlPointCount() {
        double area = parseAreaSquareMeters(areaField.getText());
        if (area <= 0d) {
            return null;
        }
        return Math.max(5, (int) Math.ceil(area / 1000d));
    }

    private Integer calculatePprPointCount() {
        double areaSquareMeters = parseAreaSquareMeters(areaField.getText());
        if (areaSquareMeters <= 0d) {
            return null;
        }
        double areaHectares = areaSquareMeters / 10000d;
        if (areaHectares <= 5d) {
            return Math.max(10, (int) Math.ceil(areaHectares * 15d));
        }
        if (areaHectares <= 10d) {
            return Math.max(75, (int) Math.ceil(areaHectares * 10d));
        }
        return Math.max(100, (int) Math.ceil(areaHectares * 5d));
    }

    private double parseAreaSquareMeters(String value) {
        if (value == null) {
            return 0d;
        }
        String normalized = value.toLowerCase(Locale.ROOT)
                .replace('\u00A0', ' ')
                .replace("кв. м", "")
                .replace("м2", "")
                .replace("м²", "")
                .replaceAll("[^0-9,.-]", "")
                .replace(",", ".");
        if (normalized.isBlank()) {
            return 0d;
        }
        try {
            return Double.parseDouble(normalized);
        } catch (NumberFormatException ex) {
            return 0d;
        }
    }

    private JPanel createImportPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 8));
        JButton btnLoad = new JButton(
                "Загрузить ТЗ",
                FontIcon.of(FontAwesomeSolid.FILE_UPLOAD, 16, Color.WHITE)
        );
        btnLoad.setFocusPainted(false);
        btnLoad.setBackground(new Color(33, 150, 243));
        btnLoad.setForeground(Color.WHITE);
        btnLoad.setFont(UIManager.getFont("Button.font"));
        btnLoad.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                btnLoad.setBackground(new Color(30, 136, 229));
            }

            @Override
            public void mouseExited(MouseEvent e) {
                btnLoad.setBackground(new Color(33, 150, 243));
            }
        });
        btnLoad.addActionListener(e -> importTechnicalAssignment());

        JButton btnExport = new JButton(
                "Экспорт протокола РАД",
                FontIcon.of(FontAwesomeSolid.FILE_EXCEL, 16, Color.WHITE)
        );
        btnExport.setFocusPainted(false);
        btnExport.setBackground(new Color(239, 108, 0));
        btnExport.setForeground(Color.WHITE);
        btnExport.setFont(UIManager.getFont("Button.font"));
        btnExport.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                btnExport.setBackground(new Color(230, 92, 0));
            }

            @Override
            public void mouseExited(MouseEvent e) {
                btnExport.setBackground(new Color(239, 108, 0));
            }
        });
        btnExport.addActionListener(e -> AreaProtocolTitleExporter.export(getProtocolData(), this));

        JButton btnExportNoise = new JButton(
                "Экспорт протокола ШУМ",
                FontIcon.of(FontAwesomeSolid.FILE_EXCEL, 16, Color.WHITE)
        );
        btnExportNoise.setFocusPainted(false);
        btnExportNoise.setBackground(new Color(239, 108, 0));
        btnExportNoise.setForeground(Color.WHITE);
        btnExportNoise.setFont(UIManager.getFont("Button.font"));
        btnExportNoise.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                btnExportNoise.setBackground(new Color(230, 92, 0));
            }

            @Override
            public void mouseExited(MouseEvent e) {
                btnExportNoise.setBackground(new Color(239, 108, 0));
            }
        });
        btnExportNoise.addActionListener(e -> AreaNoiseProtocolExporter.export(getProtocolData(), this));

        panel.add(btnLoad);
        panel.add(btnExport);
        panel.add(btnExportNoise);
        return panel;
    }

    public void requestSaveProject() {
        saveAreaProject();
    }

    public void requestLoadProject() {
        loadAreaProject();
    }

    public String getProjectName() {
        return projectNameField.getText().trim();
    }

    private void saveAreaProject() {
        String projectName = projectNameField.getText().trim();
        if (projectName.isBlank()) {
            JOptionPane.showMessageDialog(this,
                    "Введите название проекта на титульной странице Рад.",
                    "Сохранить проект",
                    JOptionPane.WARNING_MESSAGE);
            projectNameField.requestFocusInWindow();
            return;
        }

        String savedName = generateAreaProjectVersionName(projectName);
        try {
            DatabaseManager.saveAreaProject(savedName, serializeAreaProjectSnapshot(createAreaProjectSnapshot(savedName)));
            projectNameField.setText(extractAreaProjectBaseName(savedName));
            JOptionPane.showMessageDialog(this,
                    "Проект сохранен:\n" + savedName,
                    "Сохранить проект",
                    JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException | SQLException ex) {
            JOptionPane.showMessageDialog(this,
                    "Не удалось сохранить проект: " + ex.getMessage(),
                    "Сохранить проект",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void loadAreaProject() {
        try {
            List<DatabaseManager.AreaProjectInfo> projects = DatabaseManager.getAllAreaProjects();
            if (projects.isEmpty()) {
                JOptionPane.showMessageDialog(this,
                        "Нет сохраненных проектов участков.",
                        "Загрузить проект",
                        JOptionPane.INFORMATION_MESSAGE);
                return;
            }
            DatabaseManager.AreaProjectInfo selectedProject = pickAreaProject(projects);
            if (selectedProject == null) {
                return;
            }
            byte[] snapshotBytes = DatabaseManager.loadAreaProjectSnapshot(selectedProject.getId());
            if (snapshotBytes == null || snapshotBytes.length == 0) {
                throw new IOException("Проект участка пустой или поврежден.");
            }
            applyAreaProjectSnapshot(deserializeAreaProjectSnapshot(snapshotBytes), selectedProject.getName());
            JOptionPane.showMessageDialog(this,
                    "Проект загружен:\n" + selectedProject.getName(),
                    "Загрузить проект",
                    JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException | ClassNotFoundException | SQLException ex) {
            JOptionPane.showMessageDialog(this,
                    "Не удалось загрузить проект: " + ex.getMessage(),
                    "Загрузить проект",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private AreaProjectSnapshot createAreaProjectSnapshot(String projectName) throws IOException {
        AreaProjectSnapshot snapshot = new AreaProjectSnapshot();
        snapshot.projectName = projectName;
        snapshot.protocolDate = getDateText(protocolDatePicker);
        snapshot.customerNameAndContacts = getFieldText(customerNameContactsField, CUSTOMER_PLACEHOLDER);
        snapshot.customerLegalAddress = customerLegalAddressField.getText().trim();
        snapshot.customerActualAddress = customerActualAddressField.getText().trim();
        snapshot.objectName = objectNameArea.getText().trim();
        snapshot.objectAddress = objectAddressField.getText().trim();
        snapshot.contractNumber = contractNumberField.getText().trim();
        snapshot.contractDate = getDateText(contractDatePicker);
        snapshot.applicationNumber = applicationNumberField.getText().trim();
        snapshot.applicationDate = getDateText(applicationDatePicker);
        snapshot.representative = representativeField.getText().trim();
        snapshot.areaText = areaField.getText().trim();
        List<MedValueSnapshot> medValues = collectMedValueSnapshots();
        snapshot.medValues = medValues;
        snapshot.gammaMinValue = gammaMinValueField.getText().trim();
        snapshot.gammaMaxValue = gammaMaxValueField.getText().trim();
        snapshot.gammaAverageValue = gammaAverageValueField.getText().trim();
        snapshot.pprMinValue = pprMinValueField.getText().trim();
        snapshot.pprMaxValue = pprMaxValueField.getText().trim();
        snapshot.noiseEquivalentMinValue = noiseEquivalentMinValueField.getText().trim();
        snapshot.noiseEquivalentMaxValue = noiseEquivalentMaxValueField.getText().trim();
        snapshot.noiseMaxLevelMinValue = noiseMaxLevelMinValueField.getText().trim();
        snapshot.noiseMaxLevelMaxValue = noiseMaxLevelMaxValueField.getText().trim();
        snapshot.noiseMethod = String.valueOf(noiseMethodComboBox.getSelectedItem());
        snapshot.medSelected = medCheckBox.isSelected();
        snapshot.pprSelected = pprCheckBox.isSelected();
        snapshot.imageName = imageFile == null ? imageLabel.getText() : imageFile.getName();
        snapshot.imageBytes = encodeSketchImage();
        snapshot.boundaryColorRgb = sketchPreviewPanel.getBoundaryColor().getRGB();
        snapshot.gammaControlPointScale = sketchPreviewPanel.getGammaControlPointScale();
        snapshot.pprPointScale = sketchPreviewPanel.getPprPointScale();
        snapshot.noisePointScale = sketchPreviewPanel.getNoisePointScale();
        snapshot.measurementRows = collectMeasurementProjectRows();
        snapshot.boundaryPoints = sketchPreviewPanel.getBoundaryPointSnapshots();
        snapshot.gammaProfiles = sketchPreviewPanel.getGammaProfileSnapshots();
        snapshot.gammaControlPoints = sketchPreviewPanel.getGammaControlPointSnapshots();
        snapshot.pprPoints = sketchPreviewPanel.getPprPointSnapshots();
        snapshot.noisePoints = sketchPreviewPanel.getNoisePointSnapshots();
        return snapshot;
    }

    private void applyAreaProjectSnapshot(AreaProjectSnapshot snapshot, String fallbackProjectName) throws IOException {
        String loadedProjectName = snapshot.projectName == null || snapshot.projectName.isBlank()
                ? fallbackProjectName
                : snapshot.projectName;
        setFieldText(projectNameField, loadedProjectName);
        setDatePickerDate(protocolDatePicker, snapshot.protocolDate);
        setFieldText(customerNameContactsField, snapshot.customerNameAndContacts);
        setFieldText(customerLegalAddressField, snapshot.customerLegalAddress);
        setFieldText(customerActualAddressField, snapshot.customerActualAddress);
        objectNameArea.setText(snapshot.objectName == null ? "" : snapshot.objectName);
        setFieldText(objectAddressField, snapshot.objectAddress);
        setFieldText(contractNumberField, snapshot.contractNumber);
        setDatePickerDate(contractDatePicker, snapshot.contractDate);
        setFieldText(applicationNumberField, snapshot.applicationNumber);
        setDatePickerDate(applicationDatePicker, snapshot.applicationDate);
        setFieldText(representativeField, snapshot.representative);
        setFieldText(areaField, snapshot.areaText);
        setFieldText(pprMinValueField, snapshot.pprMinValue);
        setFieldText(pprMaxValueField, snapshot.pprMaxValue);
        setFieldText(noiseEquivalentMinValueField, snapshot.noiseEquivalentMinValue);
        setFieldText(noiseEquivalentMaxValueField, snapshot.noiseEquivalentMaxValue);
        setFieldText(noiseMaxLevelMinValueField, snapshot.noiseMaxLevelMinValue);
        setFieldText(noiseMaxLevelMaxValueField, snapshot.noiseMaxLevelMaxValue);
        selectNoiseMethod(snapshot.noiseMethod);
        medCheckBox.setSelected(snapshot.medSelected);
        pprCheckBox.setSelected(snapshot.pprSelected);
        applyMeasurementProjectRows(snapshot.measurementRows);

        BufferedImage loadedImage = decodeSketchImage(snapshot.imageBytes);
        imageFile = null;
        imageLabel.setText(snapshot.imageName == null || snapshot.imageName.isBlank()
                ? "Картинка не выбрана"
                : snapshot.imageName);
        sketchPreviewPanel.applySketchSnapshot(snapshot, loadedImage);
        applyMedValueSnapshots(snapshot.medValues, snapshot.gammaMinValue,
                snapshot.gammaMaxValue, snapshot.gammaAverageValue);
        if (gammaControlPointSizeSlider != null) {
            gammaControlPointSizeSlider.setValue(sketchPreviewPanel.getGammaControlPointScale());
        }
        if (pprPointSizeSlider != null) {
            pprPointSizeSlider.setValue(sketchPreviewPanel.getPprPointScale());
        }
        if (noisePointSizeSlider != null) {
            noisePointSizeSlider.setValue(sketchPreviewPanel.getNoisePointScale());
        }
    }

    private void selectNoiseMethod(String method) {
        if (method == null || method.isBlank()) {
            noiseMethodComboBox.setSelectedItem(NOISE_METHOD_MI_LABEL);
            return;
        }
        if (method.equals(NOISE_METHOD_MI)) {
            method = NOISE_METHOD_MI_LABEL;
        } else if (method.equals(NOISE_METHOD_ECOFIZIKA)) {
            method = NOISE_METHOD_ECOFIZIKA_LABEL;
        }
        for (int i = 0; i < noiseMethodComboBox.getItemCount(); i++) {
            if (method.equals(noiseMethodComboBox.getItemAt(i))) {
                noiseMethodComboBox.setSelectedIndex(i);
                return;
            }
        }
        noiseMethodComboBox.setSelectedItem(NOISE_METHOD_MI_LABEL);
    }

    private byte[] serializeAreaProjectSnapshot(AreaProjectSnapshot snapshot) throws IOException {
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        try (ObjectOutputStream output = new ObjectOutputStream(bytes)) {
            output.writeObject(snapshot);
        }
        return bytes.toByteArray();
    }

    private AreaProjectSnapshot deserializeAreaProjectSnapshot(byte[] snapshotBytes)
            throws IOException, ClassNotFoundException {
        try (ObjectInputStream input = new ObjectInputStream(new ByteArrayInputStream(snapshotBytes))) {
            Object object = input.readObject();
            if (!(object instanceof AreaProjectSnapshot)) {
                throw new IOException("Данные не являются проектом участка.");
            }
            return (AreaProjectSnapshot) object;
        }
    }

    private DatabaseManager.AreaProjectInfo pickAreaProject(List<DatabaseManager.AreaProjectInfo> projects) {
        Window owner = SwingUtilities.getWindowAncestor(this);
        JDialog dialog = new JDialog(owner, "Загрузить проект участка", Dialog.ModalityType.APPLICATION_MODAL);
        dialog.setSize(520, 420);
        dialog.setLocationRelativeTo(this);
        dialog.setLayout(new BorderLayout(10, 10));

        DefaultListModel<DatabaseManager.AreaProjectInfo> model = new DefaultListModel<>();
        for (DatabaseManager.AreaProjectInfo project : projects) {
            model.addElement(project);
        }
        JList<DatabaseManager.AreaProjectInfo> list = new JList<>(model);
        list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        list.setFixedCellHeight(28);
        list.setCellRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index,
                                                          boolean isSelected, boolean cellHasFocus) {
                Component component = super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (value instanceof DatabaseManager.AreaProjectInfo) {
                    setText(((DatabaseManager.AreaProjectInfo) value).getName());
                }
                return component;
            }
        });
        dialog.add(new JScrollPane(list), BorderLayout.CENTER);

        final DatabaseManager.AreaProjectInfo[] selectedProject = new DatabaseManager.AreaProjectInfo[1];
        JButton deleteButton = new JButton("Удалить...");
        JButton okButton = new JButton("ОК");
        JButton cancelButton = new JButton("Отмена");
        deleteButton.setEnabled(false);
        okButton.setEnabled(false);

        list.addListSelectionListener(e -> {
            boolean hasSelection = !list.isSelectionEmpty();
            deleteButton.setEnabled(hasSelection);
            okButton.setEnabled(hasSelection);
        });
        list.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2 && !list.isSelectionEmpty()) {
                    selectedProject[0] = list.getSelectedValue();
                    dialog.dispose();
                }
            }
        });

        deleteButton.addActionListener(e -> deleteSelectedAreaProject(dialog, model, list));
        okButton.addActionListener(e -> {
            selectedProject[0] = list.getSelectedValue();
            dialog.dispose();
        });
        cancelButton.addActionListener(e -> dialog.dispose());

        JPanel buttons = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttons.add(deleteButton);
        buttons.add(okButton);
        buttons.add(cancelButton);
        dialog.add(buttons, BorderLayout.SOUTH);

        dialog.setVisible(true);
        return selectedProject[0];
    }

    private void deleteSelectedAreaProject(JDialog dialog,
                                           DefaultListModel<DatabaseManager.AreaProjectInfo> model,
                                           JList<DatabaseManager.AreaProjectInfo> list) {
        if (list.isSelectionEmpty()) {
            return;
        }
        DatabaseManager.AreaProjectInfo project = list.getSelectedValue();
        int answer = JOptionPane.showConfirmDialog(
                dialog,
                "Вы точно хотите удалить проект участка:\n«" + project.getName() + "»?",
                "Подтвердите удаление",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.WARNING_MESSAGE
        );
        if (answer != JOptionPane.YES_OPTION) {
            return;
        }
        try {
            DatabaseManager.deleteAreaProject(project.getId());
            int index = list.getSelectedIndex();
            model.remove(index);
            if (!model.isEmpty()) {
                list.setSelectedIndex(Math.min(index, model.getSize() - 1));
            }
            JOptionPane.showMessageDialog(dialog,
                    "Проект удален.",
                    "Готово",
                    JOptionPane.INFORMATION_MESSAGE);
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(dialog,
                    "Ошибка удаления: " + ex.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private String generateAreaProjectVersionName(String baseName) {
        String currentDate = LocalDate.now().format(DATE_FORMAT);
        String cleanBaseName = extractAreaProjectBaseName(baseName);
        int nextVersion = calculateNextAreaProjectVersion(cleanBaseName);
        return nextVersion == 1
                ? cleanBaseName
                : cleanBaseName + " ред." + nextVersion + " " + currentDate;
    }

    private int calculateNextAreaProjectVersion(String cleanBaseName) {
        try {
            Pattern versionPattern = Pattern.compile("^" + Pattern.quote(cleanBaseName)
                    + "(?: ред\\.(\\d+) (\\d{2}\\.\\d{2}\\.\\d{4}))?$");
            return DatabaseManager.getAllAreaProjects().stream()
                    .map(project -> versionPattern.matcher(project.getName()))
                    .filter(Matcher::find)
                    .mapToInt(matcher -> matcher.group(1) != null ? Integer.parseInt(matcher.group(1)) : 1)
                    .max()
                    .orElse(0) + 1;
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this,
                    "Ошибка получения списка проектов: " + ex.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE);
            return 1;
        }
    }

    private String extractAreaProjectBaseName(String name) {
        if (name == null) {
            return "";
        }
        Pattern pattern = Pattern.compile("(.+?) (?:ред\\.\\d+ \\d{2}\\.\\d{2}\\.\\d{4})$");
        Matcher matcher = pattern.matcher(name.trim());
        return matcher.find() ? matcher.group(1).trim() : name.trim();
    }

    private List<MeasurementProjectRow> collectMeasurementProjectRows() {
        List<MeasurementProjectRow> rows = new ArrayList<>();
        for (MeasurementRow row : measurementRows) {
            MeasurementProjectRow snapshotRow = new MeasurementProjectRow();
            snapshotRow.date = getDateText(row.datePicker);
            snapshotRow.tempInsideStart = row.tempInsideStart.getText().trim();
            snapshotRow.tempInsideEnd = row.tempInsideEnd.getText().trim();
            snapshotRow.tempOutsideStart = row.tempOutsideStart.getText().trim();
            snapshotRow.tempOutsideEnd = row.tempOutsideEnd.getText().trim();
            snapshotRow.gammaSelected = row.gammaCheckBox.isSelected();
            snapshotRow.pprSelected = row.pprCheckBox.isSelected();
            snapshotRow.noiseSelected = row.noiseCheckBox.isSelected();
            rows.add(snapshotRow);
        }
        return rows;
    }

    private void applyMeasurementProjectRows(List<MeasurementProjectRow> rows) {
        measurementRows.clear();
        measurementRowsPanel.removeAll();
        if (rows == null || rows.isEmpty()) {
            addMeasurementRow(null);
        } else {
            for (MeasurementProjectRow snapshotRow : rows) {
                addMeasurementRow(parseDate(snapshotRow.date));
                MeasurementRow row = measurementRows.get(measurementRows.size() - 1);
                row.tempInsideStart.setText(snapshotRow.tempInsideStart == null ? "" : snapshotRow.tempInsideStart);
                row.tempInsideEnd.setText(snapshotRow.tempInsideEnd == null ? "" : snapshotRow.tempInsideEnd);
                row.tempOutsideStart.setText(snapshotRow.tempOutsideStart == null ? "" : snapshotRow.tempOutsideStart);
                row.tempOutsideEnd.setText(snapshotRow.tempOutsideEnd == null ? "" : snapshotRow.tempOutsideEnd);
                row.gammaCheckBox.setSelected(snapshotRow.gammaSelected);
                row.pprCheckBox.setSelected(snapshotRow.pprSelected);
                row.noiseCheckBox.setSelected(snapshotRow.noiseSelected);
            }
        }
        measurementRowsPanel.revalidate();
        measurementRowsPanel.repaint();
    }

    private byte[] encodeSketchImage() throws IOException {
        BufferedImage image = sketchPreviewPanel.getImage();
        if (image == null) {
            return null;
        }
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
    }

    private BufferedImage decodeSketchImage(byte[] imageBytes) throws IOException {
        if (imageBytes == null || imageBytes.length == 0) {
            return null;
        }
        return ImageIO.read(new ByteArrayInputStream(imageBytes));
    }

    private JPanel createPlaceholderPanel(String title) {
        JPanel panel = new JPanel(new GridBagLayout());
        JLabel label = new JLabel(title);
        label.setForeground(UIManager.getColor("Label.disabledForeground"));
        panel.add(label);
        return panel;
    }

    private void addLabeledField(JPanel panel, GridBagConstraints gbc,
                                 String labelText, JComponent field) {
        gbc.gridx = 0;
        gbc.weightx = 0.0;
        gbc.fill = GridBagConstraints.NONE;
        panel.add(new JLabel(labelText), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panel.add(field, gbc);

        gbc.gridy++;
    }

    private DatePicker createDatePicker(LocalDate initialDate) {
        DatePickerSettings settings = new DatePickerSettings(new Locale("ru", "RU"));
        settings.setFormatForDatesCommonEra(DATE_PATTERN);
        settings.setAllowEmptyDates(true);
        DatePicker picker = new DatePicker(settings);
        if (initialDate != null) {
            picker.setDate(initialDate);
        }
        return picker;
    }

    private void installPlaceholder(JTextField field, String placeholder) {
        Color normalColor = field.getForeground();
        field.setForeground(Color.GRAY);
        field.setText(placeholder);

        field.addFocusListener(new FocusAdapter() {
            @Override
            public void focusGained(FocusEvent e) {
                if (field.getText().equals(placeholder)) {
                    field.setText("");
                    field.setForeground(normalColor);
                }
            }

            @Override
            public void focusLost(FocusEvent e) {
                if (field.getText().isBlank()) {
                    field.setForeground(Color.GRAY);
                    field.setText(placeholder);
                }
            }
        });
    }

    private void addMeasurementRow(LocalDate date) {
        MeasurementRow row = new MeasurementRow();
        row.panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2));

        int index = measurementRows.size() + 1;
        row.datePicker = createDatePicker(date);
        row.tempInsideStart = new JTextField(4);
        row.tempInsideEnd = new JTextField(4);
        row.tempOutsideStart = new JTextField(4);
        row.tempOutsideEnd = new JTextField(4);
        row.gammaCheckBox = new JCheckBox("Гамма");
        row.pprCheckBox = new JCheckBox("ППР");
        row.noiseCheckBox = new JCheckBox("Шум");

        row.panel.add(new JLabel("Дата " + index + ":"));
        row.panel.add(row.datePicker);
        row.panel.add(new JLabel("Внутри, начало:"));
        row.panel.add(row.tempInsideStart);
        row.panel.add(new JLabel("конец:"));
        row.panel.add(row.tempInsideEnd);
        row.panel.add(new JLabel("Улица, начало:"));
        row.panel.add(row.tempOutsideStart);
        row.panel.add(new JLabel("конец:"));
        row.panel.add(row.tempOutsideEnd);
        row.panel.add(new JLabel("Что измеряли:"));
        row.panel.add(row.gammaCheckBox);
        row.panel.add(row.pprCheckBox);
        row.panel.add(row.noiseCheckBox);

        measurementRows.add(row);
        measurementRowsPanel.add(row.panel);
        measurementRowsPanel.revalidate();
        measurementRowsPanel.repaint();
    }

    private void removeLastMeasurementRow() {
        if (measurementRows.isEmpty()) {
            return;
        }
        MeasurementRow last = measurementRows.remove(measurementRows.size() - 1);
        measurementRowsPanel.remove(last.panel);
        measurementRowsPanel.revalidate();
        measurementRowsPanel.repaint();
    }

    private void chooseImage() {
        JFileChooser chooser = new JFileChooser(resolveLastSketchDirectory());
        chooser.setDialogTitle("Добавить эскиз");
        FileNameExtensionFilter jpgFilter = new FileNameExtensionFilter("JPG (*.jpg, *.jpeg)", "jpg", "jpeg");
        chooser.setFileFilter(jpgFilter);
        chooser.addChoosableFileFilter(new FileNameExtensionFilter(
                "Изображения (*.png, *.jpg, *.jpeg, *.bmp)", "png", "jpg", "jpeg", "bmp"));
        int result = chooser.showOpenDialog(this);
        if (result != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File selectedFile = chooser.getSelectedFile();
        try {
            BufferedImage image = ImageIO.read(selectedFile);
            if (image == null) {
                throw new IOException("Файл не является изображением.");
            }
            imageFile = selectedFile;
            imageLabel.setText(imageFile.getName());
            sketchPreviewPanel.setImage(image);
            rememberSketchDirectory(selectedFile);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this,
                    "Не удалось загрузить картинку: " + ex.getMessage(),
                    "Добавить эскиз",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private File resolveLastSketchDirectory() {
        String path = PREFS.get(PREF_SKETCH_DIR, "");
        if (!path.isBlank()) {
            File directory = new File(path);
            if (directory.isDirectory()) {
                return directory;
            }
        }
        return new File(System.getProperty("user.home"));
    }

    private void rememberSketchDirectory(File selectedFile) {
        if (selectedFile == null || selectedFile.getParentFile() == null) {
            return;
        }
        PREFS.put(PREF_SKETCH_DIR, selectedFile.getParentFile().getAbsolutePath());
    }

    private void importTechnicalAssignment() {
        JFileChooser chooser = new JFileChooser(resolveTechnicalAssignmentDirectory());
        chooser.setDialogTitle("Загрузить ТЗ (Word)");
        chooser.setFileFilter(new FileNameExtensionFilter("Word document (*.docx)", "docx"));

        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File file = chooser.getSelectedFile();
        try {
            List<TitlePageImportData> dataList = TechnicalAssignmentImporter.importAllFromFile(file);
            if (dataList.isEmpty()) {
                return;
            }
            TitlePageImportData data = pickTechnicalAssignment(dataList);
            if (data == null) {
                return;
            }
            applyImportData(data);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(
                    this,
                    "Не удалось загрузить ТЗ: " + ex.getMessage(),
                    "Загрузка ТЗ",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }

    private File resolveTechnicalAssignmentDirectory() {
        File desktop = new File(System.getProperty("user.home"), "Desktop");
        File demo = new File(desktop, "демонстрация");
        if (demo.isDirectory()) {
            return demo;
        }
        return desktop.isDirectory() ? desktop : new File(System.getProperty("user.home"));
    }

    private void applyImportData(TitlePageImportData data) {
        if (data.protocolDate() != null && !data.protocolDate().isBlank()) {
            setDatePickerDate(protocolDatePicker, data.protocolDate());
        }
        if (data.customerNameAndContacts() != null && !data.customerNameAndContacts().isBlank()) {
            setFieldText(customerNameContactsField, data.customerNameAndContacts());
        }
        if (data.customerLegalAddress() != null && !data.customerLegalAddress().isBlank()) {
            setFieldText(customerLegalAddressField, data.customerLegalAddress());
        }
        if (data.customerActualAddress() != null && !data.customerActualAddress().isBlank()) {
            setFieldText(customerActualAddressField, data.customerActualAddress());
        }
        if (data.objectName() != null && !data.objectName().isBlank()) {
            objectNameArea.setText(data.objectName());
        }
        if (data.objectAddress() != null && !data.objectAddress().isBlank()) {
            setFieldText(objectAddressField, data.objectAddress());
        }
        if (data.contractNumber() != null && !data.contractNumber().isBlank()) {
            setFieldText(contractNumberField, data.contractNumber());
        }
        if (data.contractDate() != null && !data.contractDate().isBlank()) {
            setDatePickerDate(contractDatePicker, data.contractDate());
        }
        if (data.applicationNumber() != null && !data.applicationNumber().isBlank()) {
            setFieldText(applicationNumberField, data.applicationNumber());
        }
        if (data.applicationDate() != null && !data.applicationDate().isBlank()) {
            setDatePickerDate(applicationDatePicker, data.applicationDate());
        }
    }

    private TitlePageImportData pickTechnicalAssignment(List<TitlePageImportData> dataList) {
        if (dataList.size() == 1) {
            return dataList.get(0);
        }
        String[] options = buildTechnicalAssignmentOptions(dataList);
        Object selected = JOptionPane.showInputDialog(
                this,
                "Мы нашли несколько технических заданий. Какое будем заполнять?",
                "Выбор технического задания",
                JOptionPane.QUESTION_MESSAGE,
                null,
                options,
                options[0]
        );
        if (selected == null) {
            return null;
        }
        for (int i = 0; i < options.length; i++) {
            if (options[i].equals(selected)) {
                return dataList.get(i);
            }
        }
        return dataList.get(0);
    }

    private String[] buildTechnicalAssignmentOptions(List<TitlePageImportData> dataList) {
        List<String> options = new ArrayList<>();
        for (int i = 0; i < dataList.size(); i++) {
            String option = compactOption(dataList.get(i).objectName());
            if (option.isBlank()) {
                option = "Вариант " + (i + 1);
            }
            if (options.contains(option)) {
                option = option + " (вариант " + (i + 1) + ")";
            }
            options.add(option);
        }
        return options.toArray(new String[0]);
    }

    private String compactOption(String value) {
        if (value == null) {
            return "";
        }
        String text = value.replaceAll("\\s+", " ").trim();
        if (text.length() <= 120) {
            return text;
        }
        return text.substring(0, 117) + "...";
    }

    private void setDatePickerDate(DatePicker picker, String value) {
        if (picker == null) {
            return;
        }
        picker.setDate(parseDate(value));
    }

    private LocalDate parseDate(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        try {
            return LocalDate.parse(value, DATE_FORMAT);
        } catch (Exception ignore) {
            return null;
        }
    }

    private void setFieldText(JTextField field, String value) {
        field.setForeground(defaultTextColor != null ? defaultTextColor : Color.BLACK);
        field.setText(value == null ? "" : value);
    }

    private String getDateText(DatePicker picker) {
        LocalDate date = picker == null ? null : picker.getDate();
        return date == null ? "" : date.format(DATE_FORMAT);
    }

    private String getFieldText(JTextField field, String placeholder) {
        String text = field.getText().trim();
        return text.equals(placeholder) ? "" : text;
    }

    private List<MeasurementRowData> collectMeasurementRows() {
        List<MeasurementRowData> rows = new ArrayList<>();
        for (MeasurementRow row : measurementRows) {
            rows.add(new MeasurementRowData(
                    getDateText(row.datePicker),
                    row.tempInsideStart.getText().trim(),
                    row.tempInsideEnd.getText().trim(),
                    row.tempOutsideStart.getText().trim(),
                    row.tempOutsideEnd.getText().trim(),
                    row.gammaCheckBox.isSelected(),
                    row.pprCheckBox.isSelected(),
                    row.noiseCheckBox.isSelected()
            ));
        }
        return rows;
    }

    private List<MedValueData> collectMedValues() {
        List<MedValueData> values = new ArrayList<>();
        for (MedValueRow row : medValueRows) {
            values.add(new MedValueData(
                    row.profileNumber,
                    row.distanceField.getText().trim(),
                    row.countField.getText().trim()
            ));
        }
        return values;
    }

    private static void styleInput(JTextField field) {
        Dimension preferred = field.getPreferredSize();
        field.setPreferredSize(new Dimension(preferred.width, 32));
        field.setMinimumSize(new Dimension(Math.max(preferred.width, 120), 32));
        field.setMargin(new Insets(2, 6, 2, 6));
    }

    public static final class MeasurementRowData {
        private final String date;
        private final String tempInsideStart;
        private final String tempInsideEnd;
        private final String tempOutsideStart;
        private final String tempOutsideEnd;
        private final boolean gammaSelected;
        private final boolean pprSelected;
        private final boolean noiseSelected;

        MeasurementRowData(String date, String tempInsideStart, String tempInsideEnd,
                           String tempOutsideStart, String tempOutsideEnd,
                           boolean gammaSelected, boolean pprSelected, boolean noiseSelected) {
            this.date = date;
            this.tempInsideStart = tempInsideStart;
            this.tempInsideEnd = tempInsideEnd;
            this.tempOutsideStart = tempOutsideStart;
            this.tempOutsideEnd = tempOutsideEnd;
            this.gammaSelected = gammaSelected;
            this.pprSelected = pprSelected;
            this.noiseSelected = noiseSelected;
        }

        public String getDate() {
            return date;
        }

        public String getTempInsideStart() {
            return tempInsideStart;
        }

        public String getTempInsideEnd() {
            return tempInsideEnd;
        }

        public String getTempOutsideStart() {
            return tempOutsideStart;
        }

        public String getTempOutsideEnd() {
            return tempOutsideEnd;
        }

        public boolean isGammaSelected() {
            return gammaSelected;
        }

        public boolean isPprSelected() {
            return pprSelected;
        }

        public boolean isNoiseSelected() {
            return noiseSelected;
        }
    }

    public static final class MedValueData {
        private final int profileNumber;
        private final String distance;
        private final String count;

        MedValueData(int profileNumber, String distance, String count) {
            this.profileNumber = profileNumber;
            this.distance = distance;
            this.count = count;
        }

        public int getProfileNumber() {
            return profileNumber;
        }

        public String getDistance() {
            return distance;
        }

        public String getCount() {
            return count;
        }
    }

    public static final class AreaProtocolData {
        private final String protocolDate;
        private final String customerNameAndContacts;
        private final String customerLegalAddress;
        private final String customerActualAddress;
        private final String objectName;
        private final String objectAddress;
        private final String contractNumber;
        private final String contractDate;
        private final String applicationNumber;
        private final String applicationDate;
        private final String representative;
        private final List<MeasurementRowData> measurementRows;
        private final List<MedValueData> medValues;
        private final String medMinValue;
        private final String medMaxValue;
        private final String medAverageValue;
        private final String pprMinValue;
        private final String pprMaxValue;
        private final String noiseEquivalentMinValue;
        private final String noiseEquivalentMaxValue;
        private final String noiseMaxLevelMinValue;
        private final String noiseMaxLevelMaxValue;
        private final String noiseMethod;
        private final String areaText;
        private final boolean medSelected;
        private final boolean pprSelected;
        private final BufferedImage radiationSketchImage;
        private final BufferedImage noiseSketchImage;
        private final int gammaProfileMaxNumber;
        private final int gammaControlPointCount;
        private final int pprPointCount;
        private final int noisePointCount;
        private final File imageFile;

        AreaProtocolData(String protocolDate, String customerNameAndContacts,
                         String customerLegalAddress, String customerActualAddress,
                         String objectName, String objectAddress,
                         String contractNumber, String contractDate,
                         String applicationNumber, String applicationDate,
                         String representative, List<MeasurementRowData> measurementRows,
                         List<MedValueData> medValues, String medMinValue, String medMaxValue,
                         String medAverageValue, String pprMinValue, String pprMaxValue,
                         String noiseEquivalentMinValue, String noiseEquivalentMaxValue,
                         String noiseMaxLevelMinValue, String noiseMaxLevelMaxValue,
                         String noiseMethod, String areaText, boolean medSelected, boolean pprSelected,
                         BufferedImage radiationSketchImage, BufferedImage noiseSketchImage,
                         int gammaProfileMaxNumber, int gammaControlPointCount, int pprPointCount,
                         int noisePointCount, File imageFile) {
            this.protocolDate = protocolDate;
            this.customerNameAndContacts = customerNameAndContacts;
            this.customerLegalAddress = customerLegalAddress;
            this.customerActualAddress = customerActualAddress;
            this.objectName = objectName;
            this.objectAddress = objectAddress;
            this.contractNumber = contractNumber;
            this.contractDate = contractDate;
            this.applicationNumber = applicationNumber;
            this.applicationDate = applicationDate;
            this.representative = representative;
            this.measurementRows = new ArrayList<>(measurementRows);
            this.medValues = new ArrayList<>(medValues);
            this.medMinValue = medMinValue;
            this.medMaxValue = medMaxValue;
            this.medAverageValue = medAverageValue;
            this.pprMinValue = pprMinValue;
            this.pprMaxValue = pprMaxValue;
            this.noiseEquivalentMinValue = noiseEquivalentMinValue;
            this.noiseEquivalentMaxValue = noiseEquivalentMaxValue;
            this.noiseMaxLevelMinValue = noiseMaxLevelMinValue;
            this.noiseMaxLevelMaxValue = noiseMaxLevelMaxValue;
            this.noiseMethod = noiseMethod;
            this.areaText = areaText;
            this.medSelected = medSelected;
            this.pprSelected = pprSelected;
            this.radiationSketchImage = radiationSketchImage;
            this.noiseSketchImage = noiseSketchImage;
            this.gammaProfileMaxNumber = gammaProfileMaxNumber;
            this.gammaControlPointCount = gammaControlPointCount;
            this.pprPointCount = pprPointCount;
            this.noisePointCount = noisePointCount;
            this.imageFile = imageFile;
        }

        public String getProtocolDate() {
            return protocolDate;
        }

        public String getCustomerNameAndContacts() {
            return customerNameAndContacts;
        }

        public String getCustomerLegalAddress() {
            return customerLegalAddress;
        }

        public String getCustomerActualAddress() {
            return customerActualAddress;
        }

        public String getObjectName() {
            return objectName;
        }

        public String getObjectAddress() {
            return objectAddress;
        }

        public String getContractNumber() {
            return contractNumber;
        }

        public String getContractDate() {
            return contractDate;
        }

        public String getApplicationNumber() {
            return applicationNumber;
        }

        public String getApplicationDate() {
            return applicationDate;
        }

        public String getRepresentative() {
            return representative;
        }

        public List<MeasurementRowData> getMeasurementRows() {
            return new ArrayList<>(measurementRows);
        }

        public List<MedValueData> getMedValues() {
            return new ArrayList<>(medValues);
        }

        public String getMedMinValue() {
            return medMinValue;
        }

        public String getMedMaxValue() {
            return medMaxValue;
        }

        public String getMedAverageValue() {
            return medAverageValue;
        }

        public String getPprMinValue() {
            return pprMinValue;
        }

        public String getPprMaxValue() {
            return pprMaxValue;
        }

        public String getNoiseEquivalentMinValue() {
            return noiseEquivalentMinValue;
        }

        public String getNoiseEquivalentMaxValue() {
            return noiseEquivalentMaxValue;
        }

        public String getNoiseMaxLevelMinValue() {
            return noiseMaxLevelMinValue;
        }

        public String getNoiseMaxLevelMaxValue() {
            return noiseMaxLevelMaxValue;
        }

        public String getNoiseMethod() {
            return noiseMethod;
        }

        public String getAreaText() {
            return areaText;
        }

        public boolean isMedSelected() {
            return medSelected;
        }

        public boolean isPprSelected() {
            return pprSelected;
        }

        public File getImageFile() {
            return imageFile;
        }

        public BufferedImage getRadiationSketchImage() {
            return radiationSketchImage;
        }

        public BufferedImage getNoiseSketchImage() {
            return noiseSketchImage;
        }

        public int getGammaProfileMaxNumber() {
            return gammaProfileMaxNumber;
        }

        public int getGammaControlPointCount() {
            return gammaControlPointCount;
        }

        public int getPprPointCount() {
            return pprPointCount;
        }

        public int getNoisePointCount() {
            return noisePointCount;
        }
    }
}
