import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.input.MouseEvent;
import javafx.scene.paint.Color;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.util.Duration;

import java.io.File;
import java.io.IOException;

public class RootDirChoice extends javafx.application.Application{
    private double xOffset;
    private double yOffset;

    public static String DIR_PATH = "";

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(RootDirChoice.class.getResource("panes/rootDirChoice.fxml"));
        Scene scene = new Scene(fxmlLoader.load());
        scene.setFill(Color.TRANSPARENT);
        stage.initStyle(StageStyle.TRANSPARENT);
        scene.setOnMousePressed(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                xOffset = stage.getX() - event.getScreenX();
                yOffset = stage.getY() - event.getScreenY();
            }
        });
        scene.setOnMouseDragged(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                stage.setX(event.getScreenX() + xOffset);
                stage.setY(event.getScreenY() + yOffset);
            }
        });
        stage.setScene(scene);
        stage.show();
    }

    @FXML
    private Button closeButton;

    @FXML
    private Button openButton;

    @FXML
    private Label errorLabel;

    boolean checkStart = true;
    int checkStartCount = 0;

    @FXML
    void initialize() {

        Timeline timeline = new Timeline(new KeyFrame(Duration.seconds(1), e -> {
            if (!checkStart){
                if (checkStartCount == 5)
                {
                    checkStart = true;
                    openButton.setVisible(true);
                    errorLabel.setVisible(false);
                    checkStartCount=0;
                }
                checkStartCount++;
            }
        }));
        timeline.setCycleCount(-1);
        timeline.play();

        errorLabel.setVisible(false);

        openButton.setOnAction(ActionEvent -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File dir = directoryChooser.showDialog(new Stage());
            if (dir.getPath().contains("sampling")){
                Application application = new Application();
                Application.rootDirPath = dir.getPath();
                DIR_PATH = dir.getPath();
                try {
                    application.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
                Stage stage = (Stage) openButton.getScene().getWindow();
                stage.close();
            } else {
                openButton.setVisible(false);
                errorLabel.setVisible(true);
                checkStart = false;
            }
        });

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });
    }
}
