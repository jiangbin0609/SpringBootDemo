package net.intelink.excelpoi.utils;

/** 
* @ClassName: ExcelUtil 
* @Description: TODO(这里用一句话描述这个类的作用) 
* @author jiangbin
* @date 2017年12月9日 上午9:18:41 
*  
*/
public class ExcelUtil {
	public static final String OFFICE_EXCEL_2003POSTFIX = "xls";
	public static final String OFFICE_EXCEL_2010POSTFIX = "xlxs";
	public static final String EMPTY = "";
	public static final String POINT = ".";


	public static String getPostfix(String path) {
		if (path == null || EMPTY.equals(path.trim())) {
			return EMPTY;
		}
		if(path.contains(POINT)) {
			return path.substring(path.lastIndexOf(POINT)+1, path.length());
		}
		return EMPTY;
	}
}
