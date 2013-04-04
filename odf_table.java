/**
 *this code is distributed under BSD license
 * @author mohamed elimam abudrais 
 */
package com.example;

import java.util.ArrayList;

/**
 *
 * @author „Õ„œ
 */
public class odf_table {
private int col_num;
private int row_num;
public ArrayList<Cell> cells= new <Cell>ArrayList();
    public odf_table(int number_of_colomn,int number_of_row) {
        row_num=number_of_row;
        col_num=number_of_colomn;
        for(int x=0;x<col_num;x++){
            for(int y=0;y<row_num;y++){
                cells.add(new Cell());
            }
        }
    }    

    public int getCol_num() {
        return col_num;
    }

    public int getRow_num() {
        return row_num;
    }
  public void set_cell_text(int row,int col,String str){
      cells.get(row*col_num+col).text=str;
  }
   public String get_cell_text(int row,int col){
       return cells.get(row*col_num+col).text;
  }
  public void set_colomn_width(int colomn,int width){
      for(int i=0;i<row_num;i++){
          cells.get(i*col_num+colomn).width=width;
      }
  }
  public void set_row_background_colour(int row,int r,int g,int b){
      for(int i=0;i<col_num;i++){
          cells.get(i*col_num+row).r=r;
          cells.get(i*col_num+row).g=g;
          cells.get(i*col_num+row).b=b;
      }
  }
  public String get_cell_name(int row,int colomn){
      String letter="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        String result= letter.charAt(colomn) +new Integer(row+1).toString();
        return result;
  }
  public void set_cell_formula(int row,int col,String str){
      
  }
  private class Cell {
      public int width;
      public String text;
      public int r;
      public int g;
      public int b;
      public String formula;
  }
}
