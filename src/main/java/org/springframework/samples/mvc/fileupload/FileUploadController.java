package org.springframework.samples.mvc.fileupload;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mvc.extensions.ajax.AjaxUtils;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.context.request.WebRequest;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/fileupload")
public class FileUploadController {

	@ModelAttribute
	public void ajaxAttribute(WebRequest request, Model model) {
		model.addAttribute("ajaxRequest", AjaxUtils.isAjaxRequest(request));
	}

	@RequestMapping(method=RequestMethod.GET)
	public void fileUploadForm() {
	}

	@RequestMapping(method=RequestMethod.POST)
	public void processUpload(@RequestParam MultipartFile file, Model model) throws IOException {
		String name=file.getName();
		String originaFilename=file.getOriginalFilename();
		Workbook wb = readExcel(file);
		List<Map<String,String>> list = null;
		Sheet sheet = null;
		Row row = null;
		String cellData = null;
		String columns[] = {"name","age","score"};
		if(wb != null){
			//用来存放表中数据
			list = new ArrayList<>();
			//获取第一个sheet
			sheet = wb.getSheetAt(0);
			//获取最大行数
			int rownum = sheet.getPhysicalNumberOfRows();
			//获取第一行
			row = sheet.getRow(0);
			//获取最大列数
			int colnum = row.getPhysicalNumberOfCells();
			for (int i = 1; i<rownum; i++) {
				Map<String,String> map = new LinkedHashMap<>();
				row = sheet.getRow(i);
				if(row !=null){
					for (int j=0;j<colnum;j++){
						cellData = (String) getCellFormatValue(row.getCell(j));
						map.put(columns[j], cellData);
					}
				}else{
					break;
				}
				list.add(map);
			}
		}
		//遍历解析出来的list
		for (Map<String,String> map : list) {
			for (Map.Entry<String,String> entry : map.entrySet()) {
				System.out.print(entry.getKey()+":"+entry.getValue()+",");
			}
			System.out.println();
		}

		model.addAttribute("message", "File '" + file.getOriginalFilename() + "' uploaded successfully");
	}


	//读取excel
	public static Workbook readExcel(MultipartFile file){
		String filePath=file.getOriginalFilename();
		Workbook wb = null;
		if(filePath==null){
			return null;
		}
		String extString = filePath.substring(filePath.lastIndexOf("."));
		InputStream is = null;
		try {
			is=file.getInputStream();
			if(".xls".equals(extString)){
				return wb = new HSSFWorkbook(is);
			}else if(".xlsx".equals(extString)){
				return wb = new XSSFWorkbook(is);
			}else{
				return wb = null;
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return wb;
	}

	public static Object getCellFormatValue(Cell cell){
		Object cellValue = null;
		DecimalFormat df = new DecimalFormat("0");
		if(cell!=null){
			//判断cell类型
			switch(cell.getCellType()){
				case Cell.CELL_TYPE_NUMERIC:{
					cellValue = df.format(cell.getNumericCellValue());
					break;
				}
				case Cell.CELL_TYPE_FORMULA:{
					//判断cell是否为日期格式
					if(DateUtil.isCellDateFormatted(cell)){
						//转换为日期格式YYYY-mm-dd
						cellValue = cell.getDateCellValue();
					}else{
						//数字
						cellValue = df.format(cell.getNumericCellValue());
					}
					break;
				}
				case Cell.CELL_TYPE_STRING:{
					cellValue = cell.getRichStringCellValue().getString();
					break;
				}
				default:
					cellValue = "";
			}
		}else{
			cellValue = "";
		}
		return cellValue;
	}





}
