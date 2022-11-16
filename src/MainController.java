import data.BacteriasInfo;
import data.InfoList;
import fileView.XLXSOpen;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.shape.Circle;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.ResourceBundle;

public class MainController {

    class EditingCell extends TableCell<BacteriasInfo, String> {

        private TextField textField;

        public EditingCell() {
        }

        @Override
        public void startEdit() {
            if (!isEmpty()) {
                super.startEdit();
                createTextField();
                setText(null);
                setGraphic(textField);
                textField.selectAll();
            }
        }

        @Override
        public void cancelEdit() {
            super.cancelEdit();

            setText((String) getItem());
            setGraphic(null);
        }

        @Override
        public void updateItem(String item, boolean empty) {
            super.updateItem(item, empty);

            if (empty) {
                setText(null);
                setGraphic(null);
            } else {
                if (isEditing()) {
                    if (textField != null) {
                        textField.setText(getString());
                    }
                    setText(null);
                    setGraphic(textField);
                } else {
                    setText(getString());
                    setGraphic(null);
                }
            }
        }

        private void createTextField() {
            textField = new TextField(getString());
            textField.setMinWidth(this.getWidth() - this.getGraphicTextGap()* 2);
            textField.focusedProperty().addListener(new ChangeListener<Boolean>(){
                @Override
                public void changed(ObservableValue<? extends Boolean> arg0,
                                    Boolean arg1, Boolean arg2) {
                    if (!arg2) {
                        commitEdit(textField.getText());
                    }
                }
            });
        }

        private String getString() {
            return getItem() == null ? "" : getItem().toString();
        }
    }

    public InfoList infoList;
    ArrayList<String> content_list = new ArrayList<>();
    List<File> samplePath;
    MainLoader docLoad;
    XLXSOpen xlxsOpen;
    File saveSample;
    boolean checkLoad, checkUnload, checkStart = false;
    int counter, counter_files;
    public static String errorMessageStr = "";

    @FXML
    private ResourceBundle resources;

    @FXML
    private URL location;

    @FXML
    private Button dirLoadButton;

    @FXML
    private Button dirUnloadButton;

    @FXML
    private Text loadStatus;

    @FXML
    private Text loadStatus_end;

    @FXML
    private Text loadStatusFileNumber;

    @FXML
    private Button startButton;

    @FXML
    public Label lowLoadText = new Label("");

    @FXML
    public Button closeButton;

    @FXML
    private TableView<BacteriasInfo> bacterViewTable;

    @FXML
    private Button addBacterInTable;

    @FXML
    private Button removeBacterInTable;

    public MainController() throws IOException, InvalidFormatException {
    }

    int getCounter(int rowCount, int currentNumber) {
        Double temp = new Double(100/rowCount);
        return temp.intValue() + currentNumber;
    }

    boolean maleSample = false;
    boolean femaleSample = false;

    public void addHinds(){

        Tooltip tipLoad = new Tooltip();
        tipLoad.setText("Выберите папку, в которой находятся xlsx файлы");
        tipLoad.setStyle("-fx-text-fill: turquoise;");
        dirLoadButton.setTooltip(tipLoad);

        Tooltip tipUnLoad = new Tooltip();
        tipUnLoad.setText("Выберите файл, в который необходимо вставить данные по образцам");
        tipUnLoad.setStyle("-fx-text-fill: turquoise;");
        dirUnloadButton.setTooltip(tipUnLoad);

        Tooltip tipStart = new Tooltip();
        tipStart.setText("Нажмите, для того, чтобы получить выборку");
        tipStart.setStyle("-fx-text-fill: turquoise;");
        startButton.setTooltip(tipStart);

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть приложение");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

    }

    public void removeHinds(){
        dirLoadButton.setTooltip(null);
        dirUnloadButton.setTooltip(null);
        startButton.setTooltip(null);
        closeButton.setTooltip(null);
    }

    public static boolean tempHints = true;

    SerBacterList serList = new SerBacterList();

    public void setTableBacter() throws IOException {
        FileOutputStream outputStream = new FileOutputStream(Application.rootDirPath + "\\tableBacter.ser");
        ObjectOutputStream objectOutputStream = new ObjectOutputStream(outputStream);
        objectOutputStream.writeObject(serList);
        objectOutputStream.close();
    }

    public void getTableBacter() throws IOException, ClassNotFoundException {
        FileInputStream fileInputStream = new FileInputStream( Application.rootDirPath + "\\tableBacter.ser");
        ObjectInputStream objectInputStream = new ObjectInputStream(fileInputStream);
        serList = (SerBacterList) objectInputStream.readObject();
    }

    @FXML
    void initialize() throws IOException, InterruptedException, ClassNotFoundException {
        addHinds();
        FileInputStream loadStream = new FileInputStream(Application.rootDirPath + "\\load.png");
        Image loadImage = new Image(loadStream);
        ImageView loadView = new ImageView(loadImage);
        dirLoadButton.graphicProperty().setValue(loadView);

        FileInputStream unloadStream = new FileInputStream(Application.rootDirPath + "\\unload.png");
        Image unloadImage = new Image(unloadStream);
        ImageView unloadView = new ImageView(unloadImage);
        dirUnloadButton.graphicProperty().setValue(unloadView);

        FileInputStream startStream = new FileInputStream(Application.rootDirPath + "\\start.png");
        Image startImage = new Image(startStream);
        ImageView startView = new ImageView(startImage);
        startButton.graphicProperty().setValue(startView);

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);


        int r = 60;
        startButton.setShape(new Circle(r));
        startButton.setMinSize(r*2, r*2);
        startButton.setMaxSize(r*2, r*2);

        checkLoad = false;
        checkUnload = false;

        TableColumn bacteriaColumn = new TableColumn("Бактерия");
        bacteriaColumn.setCellValueFactory(new PropertyValueFactory<BacteriasInfo, String>("bacteriaName"));
        bacterViewTable.getColumns().add(bacteriaColumn);
        bacterViewTable.setEditable(true);
        Callback<TableColumn, TableCell> cellFactory =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };
        bacteriaColumn.setPrefWidth(198);
        bacteriaColumn.setCellFactory(cellFactory);
        try {
            getTableBacter();
        } catch (IOException e) {
            e.printStackTrace();
        }
        for(int i = 0; i< serList.serBacterList.size(); i++)
        {
            BacteriasInfo bacteriasInfo = new BacteriasInfo();
            bacteriasInfo.setBacteriaName(serList.serBacterList.get(i));
            bacterViewTable.getItems().add(bacteriasInfo);
        }
        bacteriaColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<BacteriasInfo, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<BacteriasInfo, String> t) {
                        serList.serBacterList.set(t.getTablePosition().getRow(), t.getNewValue());
                        bacterViewTable.getItems().get(t.getTablePosition().getRow()).setBacteriaName(t.getNewValue());
                    }
                }
        );

        removeBacterInTable.setOnAction(ActionEvent -> {
            if (bacterViewTable.getSelectionModel().getSelectedIndex() == -1)
            {
                MainController.errorMessageStr = "Вы не выбрали строку для удаления";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            else{
                serList.serBacterList.remove(bacterViewTable.getSelectionModel().getSelectedIndex());
                bacterViewTable.getItems().remove(bacterViewTable.getSelectionModel().getSelectedIndex());
            }
        });

        addBacterInTable.setOnAction(ActionEvent -> {;
            if (bacterViewTable.getSelectionModel().getSelectedIndex() == -1)
            {
                MainController.errorMessageStr = "Вы не выбрали строку, после которой должна вставиться новая";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            else{
                serList.serBacterList.add(bacterViewTable.getSelectionModel().getSelectedIndex()+1, "");
                bacterViewTable.getItems().add(bacterViewTable.getSelectionModel().getSelectedIndex()+1, new BacteriasInfo());
            }
        });

        closeButton.setOnAction(actionEvent -> {
            try {
                setTableBacter();
            } catch (IOException e) {
                e.printStackTrace();
            }
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        dirLoadButton.setOnAction(actionEvent -> {
            if(!checkStart)
            {
                loadStatus.setText("");
                loadStatus_end.setText("");
                loadStatusFileNumber.setText("");
                DirectoryChooser directoryChooser = new DirectoryChooser();
                File dir = directoryChooser.showDialog(new Stage());
                File[] file = dir.listFiles();
                samplePath = Arrays.asList(file);
                checkLoad = true;
            }
            else
            {
                errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });

        dirUnloadButton.setOnAction(actionEvent -> {
                    if(!checkStart)
                    {
                        loadStatus.setText("");
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        FileChooser fileChooser = new FileChooser();
                        saveSample = fileChooser.showOpenDialog(new Stage());
                        checkUnload = true;
                    }
                    else
                    {
                        errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
        startButton.setOnAction(actionEvent -> {
                    if(!checkStart){
                        loadStatus.setText("");
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        if(checkLoad & checkUnload){
                                if(samplePath.size() != 0)
                                {
                                    checkStart = true;
                                    new Thread(){
                                        @Override
                                        public void run(){
                                            counter_files = 0;
                                            MainLoader mainLoader = null;
                                            AlgOpen alg = null;
                                            try {
                                                mainLoader = new MainLoader(saveSample);
                                            } catch (IOException e) {
                                                e.printStackTrace();
                                            } catch (InvalidFormatException e) {
                                                e.printStackTrace();
                                            }
                                            for (int i = 0; i<samplePath.size();i++)
                                            {
                                                if(samplePath.get(i).getPath().contains(".xlsx"))
                                                {
                                                    loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                    counter = 0;
                                                    infoList = new InfoList(serList.serBacterList);
                                                    try {
                                                        alg = new AlgOpen(infoList);
                                                        xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                    } catch (IOException e) {
                                                        e.printStackTrace();
                                                    } catch (InvalidFormatException e) {
                                                        e.printStackTrace();
                                                    }
                                                    try {
                                                        xlxsOpen.getFileName(infoList);
                                                        xlxsOpen.getMainInfo(infoList);
                                                        xlxsOpen.getBacterInfo(infoList);
                                                        mainLoader.setAllInfo(infoList);
                                                        mainLoader.saveFile(saveSample);
                                                    } catch (IOException e) {
                                                        e.printStackTrace();
                                                    }
                                                    counter_files++;
                                                }
                                            }
                                            loadStatusFileNumber.setText("");
                                            loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                            checkStart = false;
                                        }
                                    }.start();
                                } else
                                {
                                    errorMessageStr = "Выбранная папка загрузки является пустой...";
                                    ErrorController errorController = new ErrorController();
                                    try {
                                        errorController.start(new Stage());
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    }
                            }
                        } else {
                            errorMessageStr = "Вы не указаали директорию загрузки или директорию выгрузки...";
                            ErrorController errorController = new ErrorController();
                            try {
                                errorController.start(new Stage());
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    } else
                    {
                        errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
    }
}
