package gainMiles;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadInsurance {

	public static void main(String[] args) {
		List<Data> dataList = read("Q1.xlsx");
		printData(dataList);
	}

	/**
	 * 列印資料
	 * 
	 * @param dataList
	 */
	private static void printData(List<Data> dataList) {
		System.out.println("[Benefit          , Coverage   , Category             , Plan Name, Coverage Value]");
		System.out.println("==================================================================================");
		for (Data data : dataList) {
			System.out.println(data.toString());
		}
	}

	/**
	 * 讀取檔案
	 * 
	 * @param string
	 */
	private static List<Data> read(String filepath) {
		List<Data> dataList = new ArrayList<Data>();
		try (FileInputStream fis = new FileInputStream(filepath); XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

			XSSFSheet inputSheet = workbook.getSheetAt(0);

			// 用來計算 plan 數量
			int lastCellNum = 0;
			ArrayList<String> planList = new ArrayList<String>();
			Data dataBean = new Data();
			// 取資料
			for (int i = 0; i < inputSheet.getLastRowNum(); i++) {
				// 取 plan name
				if (i == 0) {
					XSSFRow row0 = inputSheet.getRow(i);
					lastCellNum = row0.getLastCellNum();
					for (int j = 5; j < lastCellNum; j++) {
						planList.add(row0.getCell(j).getStringCellValue());
					}
					continue;
				}

				if (inputSheet.getRow(i) == null) {
					continue;
				}
				XSSFCell cell0 = inputSheet.getRow(i).getCell(0);
				XSSFCell cell1 = inputSheet.getRow(i).getCell(1);
				if (cell0 != null) {
					if (cell1 == null) {// benefit
						dataBean.setBenefit(cell0.getStringCellValue());
					} else {// coverage & data
						dataBean.setCoverage(cell0.getStringCellValue());
						dataBean.setCategory(cell1.getStringCellValue());
						for (int j = 5; j < lastCellNum; j++) {
							dataBean.setPlanName(planList.get(j - 5));
							dataBean.setCoverageValue(getCell2String(inputSheet.getRow(i).getCell(j)));
							dataList.add((Data) dataBean.clone());
						}
					}
				} else {
					if (cell1 != null) {// data
						dataBean.setCategory(cell1.getStringCellValue());
						for (int j = 5; j < lastCellNum; j++) {
							dataBean.setPlanName(planList.get(j - 5));
							dataBean.setCoverageValue(getCell2String(inputSheet.getRow(i).getCell(j)));
							dataList.add((Data) dataBean.clone());
						}
					}
				}
			}
		} catch (IOException e) {
			System.err.println("發生錯誤: " + e.getMessage());
			e.printStackTrace();
		} catch (CloneNotSupportedException e) {
			e.printStackTrace();
		}
		return dataList;
	}

	/**
	 * 取 cell 成字串
	 * @param cell
	 * @return
	 */
	private static String getCell2String(XSSFCell cell) {
		DataFormatter dataFormatter = new DataFormatter();
		switch (cell.getCellType()) {
		case NUMERIC:
			return dataFormatter.formatCellValue(cell);
		case STRING:
			return cell.getStringCellValue();
		default:
			break;
		}
		return null;
	}

	/**
	 * 資料 Bean
	 */
	static class Data implements Cloneable {
		private String benefit;
		private String coverage;
		private String category;
		private String planName;
		private String coverageValue;

		public String getBenefit() {
			return benefit;
		}

		public void setBenefit(String benefit) {
			this.benefit = benefit;
		}

		public String getCoverage() {
			return coverage;
		}

		public void setCoverage(String coverage) {
			this.coverage = coverage;
		}

		public String getCategory() {
			return category;
		}

		public void setCategory(String category) {
			this.category = category;
		}

		public String getPlanName() {
			return planName;
		}

		public void setPlanName(String planName) {
			this.planName = planName;
		}

		public String getCoverageValue() {
			return coverageValue;
		}

		public void setCoverageValue(String coverageValue) {
			this.coverageValue = coverageValue;
		}

		@Override
		protected Object clone() throws CloneNotSupportedException {
			return super.clone();
		}

		@Override
		public String toString() {
			return benefit + "," + coverage + "," + category + "," + planName + "," + coverageValue;
		}
	}
}
