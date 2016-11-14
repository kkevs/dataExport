package controller;

import java.io.Serializable;
import java.util.ArrayList;

import javax.faces.bean.ManagedBean;
import javax.faces.bean.ViewScoped;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

import model.Person;

@ViewScoped
@ManagedBean(name = "manager")
public class PropertyManager implements Serializable {
	private static final long serialVersionUID = 1L;

	ArrayList<Person> cacheList = new ArrayList<Person>();
	/*
	 * private String name; private String surname; private int age; private
	 * String city; // Getters/Setters omitted for brevity
	 */
	private Person person1 = new Person();

	public void save() {
		
		cacheList.add(person1);
	}

	public Person getPerson1() {
		return person1;
	}

	public void setPerson1(Person person1) {
		this.person1 = person1;
	}

	public void clear() {
		cacheList.clear();
	}

	public ArrayList<Person> getCacheList() {
		return cacheList;
	}

	public void setCacheList(ArrayList<Person> cacheList) {
		this.cacheList = cacheList;
	}

	public void postProcessXLS(Object document) {
		HSSFWorkbook wb = (HSSFWorkbook) document;
		HSSFSheet sheet = wb.getSheetAt(0);
		CellStyle style = wb.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());

		for (Row row : sheet) {
			for (Cell cell : row) {
				cell.setCellValue(cell.getStringCellValue().toUpperCase());
				cell.setCellStyle(style);
			}
		}
	}

}