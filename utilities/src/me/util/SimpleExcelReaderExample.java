package me.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.Gson;

import gov.diet.bean.Employee;
import gov.diet.bean.Equipment;
import gov.diet.bean.Facility;
import gov.diet.bean.LinkageSupportor;
import me.jnvc.alumni.bean.AlumniMember;

/**
 * A dirty simple program that reads an Excel file.
 * 
 * @author www.codejava.net
 *
 */
public class SimpleExcelReaderExample {

	public static void main(String[] args) throws IOException {

		SimpleExcelReaderExample excelReader = new SimpleExcelReaderExample();
		String excelFilePath = "jnvc_data.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

		Workbook workbook = new XSSFWorkbook(inputStream);
		// excelReader.readStaffSheet(staffSheet);
		// excelReader.readInfraSheet(workbook.getSheetAt(1));
		//excelReader.readEquipmentSheet(workbook.getSheetAt(0));
		excelReader.readLinkagesSupportorSheet(workbook.getSheetAt(0));
		excelReader.readJnvcSheet(workbook.getSheetAt(0));
		inputStream.close();
	}

	private void readJnvcSheet(Sheet jnvcSheet) {
		Iterator<Row> iterator = jnvcSheet.iterator();
		List<AlumniMember> alumniMemberList = new ArrayList<>();
		int cellCount;
		AlumniMember alumniMember = new AlumniMember();
		while (iterator.hasNext()) {

			cellCount = 1;

			alumniMember = new AlumniMember();
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					switch (cellCount) {
					// If the facility is available, indicate/give number
					// Adequacy of facility Is the linkageSupportor functional?
					// Provision for maintenance

					case 1:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						alumniMember.setRole(cell.getStringCellValue().trim());
						break;
					case 2:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						alumniMember.setName(cell.getStringCellValue().trim());
						break;
					case 3:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						alumniMember.setEmail(cell.getStringCellValue().trim());
						break;
					case 4:
						/*System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						alumniMember.setContact(cell.getNumericCellValue());*/
						break;
					default:
						break;

					}
					// System.out.print(cell.getStringCellValue());
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					switch (cellCount) {
					case 4:
						//System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						alumniMember.setContact(cell.getNumericCellValue());
						break;
					}
					System.out.print(cell.getNumericCellValue());
					break;
				// case 3:
				// employee.setAppointmentDate(cell.getStringCellValue());
				// break;
				default:
					System.out.print(cell.getCellType());
					break;
				}
				System.out.print(" - ");

				cellCount++;
			}
			alumniMemberList.add(alumniMember);
			for (AlumniMember member : alumniMemberList){
				System.out.println(member.getEmail());
			}
			System.out.println();

		}
		System.out.println(new Gson().toJson(alumniMemberList));
		// ((FileInputStream) workbook).close();

	
		
	}

	private void readLinkagesSupportorSheet(Sheet linkageSupportorSheet) {

		Iterator<Row> iterator = linkageSupportorSheet.iterator();
		List<LinkageSupportor> linkagesList = new ArrayList<>();
		int cellCount;
		LinkageSupportor linkageSupportor = new LinkageSupportor();
		while (iterator.hasNext()) {

			cellCount = 1;

			linkageSupportor = new LinkageSupportor();
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					switch (cellCount) {
					// If the facility is available, indicate/give number
					// Adequacy of facility Is the linkageSupportor functional?
					// Provision for maintenance

					case 1:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						linkageSupportor.setNameSupportingInstitute(cell.getStringCellValue());
						break;
					case 2:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						linkageSupportor.setKindOfSupport(cell.getStringCellValue());
						break;
					case 3:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						linkageSupportor.setTargetProgramme(cell.getStringCellValue());
						break;

					default:
						break;

					}
					// System.out.print(cell.getStringCellValue());
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					switch (cellCount) {
					case 4:
						System.out.print("(count" + cellCount + ") " + cell.getNumericCellValue());
						linkageSupportor.setTargetProgramme(""+cell.getNumericCellValue());
						break;
					}
					System.out.print(cell.getNumericCellValue());
					break;
				// case 3:
				// employee.setAppointmentDate(cell.getStringCellValue());
				// break;
				default:
					System.out.print(cell.getCellType());
					break;
				}
				System.out.print(" - ");

				cellCount++;
			}
			linkagesList.add(linkageSupportor);
			System.out.println();

		}
		System.out.println(new Gson().toJson(linkagesList));
		// ((FileInputStream) workbook).close();

	}

	private void readEquipmentSheet(Sheet equipmentSheet) {

		Iterator<Row> iterator = equipmentSheet.iterator();
		List<Equipment> equipmentList = new ArrayList<>();
		int cellCount;
		Equipment equipment = new Equipment();
		while (iterator.hasNext()) {

			cellCount = 1;

			equipment = new Equipment();
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					switch (cellCount) {
					// If the facility is available, indicate/give number
					// Adequacy of facility Is the equipment functional?
					// Provision for maintenance

					case 1:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						equipment.setName(cell.getStringCellValue());
						break;
					case 2:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						equipment.setTotalNoOfEquipments(cell.getStringCellValue());
						break;
					case 3:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						equipment.setAdequacy(cell.getStringCellValue());
						break;
					case 4:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						equipment.setEquipmentFunctional(cell.getStringCellValue());
						break;
					case 5:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						equipment.setProvisionForMaintenance(cell.getStringCellValue());
						break;

					default:
						break;

					}
					// System.out.print(cell.getStringCellValue());
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					switch (cellCount) {
					case 4:
						System.out.print("(count" + cellCount + ") " + cell.getNumericCellValue());
						break;
					}
					System.out.print(cell.getNumericCellValue());
					break;
				// case 3:
				// employee.setAppointmentDate(cell.getStringCellValue());
				// break;
				default:
					System.out.print(cell.getCellType());
					break;
				}
				System.out.print(" - ");

				cellCount++;
			}
			equipmentList.add(equipment);
			System.out.println();

		}
		System.out.println(new Gson().toJson(equipmentList));
		// ((FileInputStream) workbook).close();

	}

	private void readStaffSheet(Sheet staffSheet) throws IOException {

		Iterator<Row> iterator = staffSheet.iterator();
		List<Employee> employeeList = new ArrayList<>();
		int count = 1;
		int cellCount;
		Employee employee = new Employee();
		while (iterator.hasNext()) {

			cellCount = 1;
			if (count % 2 == 1) {
				employee = new Employee();
				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						switch (cellCount) {

						case 1:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
							employee.setName(cell.getStringCellValue());

							break;
						case 2:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
							employee.setDesignation(cell.getStringCellValue());

							break;
						case 3:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
							employee.setQualification(cell.getStringCellValue());

							break;
						case 4:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
							employee.setAppointmentDate(cell.getStringCellValue());

							break;
						case 5:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
							employee.setNatureOfAppointment(cell.getStringCellValue());

							break;

						default:
							break;

						}
						// System.out.print(cell.getStringCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						switch (cellCount) {
						case 4:
							System.out.print("(count" + cellCount + ") " + cell.getNumericCellValue());
							employee.setAppointmentDate(((Double) cell.getNumericCellValue()).toString());
							break;
						}
						System.out.print(cell.getNumericCellValue());
						break;
					// case 3:
					// employee.setAppointmentDate(cell.getStringCellValue());
					// break;
					default:
						System.out.print(cell.getCellType());
						break;
					}
					System.out.print(" - ");

					cellCount++;
				}
				count++;
			} else {
				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					/*
					 * (count1) GANGAPPA - (count3) PRINCIPAL - (count4) MA
					 * M.ED. - (count5) 24-10-1984 - (count6) PERMANENT - -
					 * (count1) (DDPI DEVELOPMENT) - - (count3) 23-10-2013
					 */
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						switch (cellCount) {
						case 1:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());

							break;
						case 2:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
							employee.setDesignation(employee.getDesignation() + " " + cell.getStringCellValue());
							break;
						case 3:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());

							break;
						case 4:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
							employee.setDateOfPosting(cell.getStringCellValue());

							break;
						case 5:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());

							break;
						case 6:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());

							break;
						case 7:
							System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());

							break;
						default:
							break;

						}
						// System.out.print(cell.getStringCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					// case 3:
					// employee.setDateOfPosting(cell.getStringCellValue());
					// break;
					default:
						System.out.print(cell.getCellType());

					}
					System.out.print(" - ");

					cellCount++;
				}
				count++;
				employeeList.add(employee);
			}
			System.out.println("---------------------------");

		}
		System.out.println(new Gson().toJson(employeeList));
		// ((FileInputStream) workbook).close();

	}

	private void readInfraSheet(Sheet infraSheet) throws IOException {

		Iterator<Row> iterator = infraSheet.iterator();
		List<Facility> facilityList = new ArrayList<>();
		int cellCount;
		Facility facility = new Facility();
		iterator.next();
		iterator.next();
		while (iterator.hasNext()) {
			cellCount = 1;
			facility = new Facility();
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					switch (cellCount) {

					case 1:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						// employee.setName(cell.getStringCellValue());
						facility.setName(cell.getStringCellValue());

						break;
					case 2:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						facility.setAvailablity(cell.getStringCellValue());
						break;
					case 3:
						System.out.print("(count" + cellCount + ") " + cell.getStringCellValue());
						facility.setAdequacy(cell.getStringCellValue());
						break;
					default:
						break;

					}
					// System.out.print(cell.getStringCellValue());
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					switch (cellCount) {
					}
					System.out.print(cell.getNumericCellValue());
					break;
				// case 3:
				// employee.setAppointmentDate(cell.getStringCellValue());
				// break;
				default:
					System.out.print(cell.getCellType());
					break;
				}
				System.out.print(" - ");

				cellCount++;
			}
			facilityList.add(facility);

		}
		System.out.println();
		System.out.println();
		System.out.println();
		System.out.println(new Gson().toJson(facilityList));
		// ((FileInputStream) workbook).close();

	}
}