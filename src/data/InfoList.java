package data;

import java.util.ArrayList;

public class InfoList {
    public String fileName = "";

    public ArrayList<String> mainInfo = new ArrayList<>();
    public ArrayList<ArrayList<String>> bacterInfo = new ArrayList<>();
    public ArrayList<ArrayList<String>> algs = new ArrayList<>();

    public InfoList(ArrayList<String> bacterNames){
        for (int i = 0; i < bacterNames.size(); i++) {
            bacterInfo.add(new ArrayList<>());
            bacterInfo.get(i).add(bacterNames.get(i));
        }
    }
}
