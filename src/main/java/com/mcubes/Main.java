package com.mcubes;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.mcubes.writer.ExcelWriter;

public class Main {

    public static void main(String[] args) {

        Main.writeSimpleData();

    }

    public static void writeSimpleData(){
        ExcelWriter writer = new ExcelWriter("files", "test_excel_01.xlsx", "Sheet_no_01");
        writer.setHeaders("Bangladesh", "India", "Island", "Germany", "Brazil");
        List<Integer[]> data = new ArrayList<>();
        data.add(new Integer[]{1,2,3,4,5});
        data.add(new Integer[]{11,22,4,56,87});
        data.add(new Integer[]{15,62,37,48,5});
        data.add(new Integer[]{5,7,8,23,7});
        data.add(new Integer[]{61,82,13,54,57});
        data.add(new Integer[]{15,25,53,74,55});
        writer.setIntegerData(data);
        try {
            writer.write();
            System.out.println("Data write success!");
        }catch (IOException e){
            e.printStackTrace();
        }
    }
}



