package com.free.modules.population.service.base;

import com.free.api.service.TokenService;
import com.free.basis.pml.DaoPlus;
import com.free.basis.pml.HashDataPlus;
import com.free.basis.pml.ListPlus;
import com.free.basis.pml.ProgressBar;
import com.free.basis.util.IdCardCheck;
import com.free.modules.population.enums.Dictionarys;
import com.free.modules.population.model.Domicile;
import com.free.modules.population.model.Person;
import com.free.modules.population.service.ImportValidation;
import com.free.modules.population.service.PersonService;
import com.free.modules.population.service.UserRightService;
import com.free.modules.util.GetInfo;
import com.free.plugins.area.AreaService;
import com.free.system.model.Depart;
import com.free.system.model.Dictionary;
import com.free.system.model.User;
import com.free.system.service.DictionaryService;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.framework.core.dao.DaoApi;
import org.framework.core.file.upload.Attachment;
import org.framework.core.file.upload.AttachmentService;
import org.framework.core.privilege.particle.AuthorizeService;
import org.framework.core.utils.ExtendFactory;
import org.framework.core.utils.HashData;
import org.springframework.beans.factory.annotation.Autowired;

import javax.servlet.http.HttpServletRequest;
import java.beans.PropertyDescriptor;
import java.io.ByteArrayInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

@SuppressWarnings("ALL")
public class BaseService {
    @Autowired
    protected DaoApi dao;
    @Autowired
    protected ExtendFactory ext;
    @Autowired
    public AttachmentService attachmentService;
    @Autowired
    public AreaService areaService;
    @Autowired
    public PersonService pservice;
    @Autowired
    protected AuthorizeService authorizeService;
    @Autowired
    protected DaoPlus daoPlus;
    @Autowired
    protected UserRightService userRightService;
    @Autowired
    protected GetInfo getCardInfo;
    @Autowired
    protected DictionaryService dic;
    @Autowired
    protected TokenService tokenService;

    protected Logger log = Logger.getLogger(BaseService.class);

    protected Dictionarys enumdic;


    public List<Object> addImage(List<Object> list){
        for (Object obj:list){
            Map<String,Object> map = (Map<String,Object>) obj;
            map.put("files",attachmentService.findFilesById(map.get("id").toString()));
        }
        return list;
    }


    /**
     * 字典获取
     * @param parentId
     * @param keyword
     * @return
     */
    public String getValue(String parentId,String keyword){
        List<Dictionary> dicList=dao.queryHql ("from Dictionary where parentId=?0 and key=?1",parentId,keyword);
        return dicList==null ? keyword :dicList.get (0).getValue ();
    }
    /**
     * 街道获取
     * @param parentId
     * @param keyword
     * @return
     */
    public String getSteer(String name){
        List<Depart> departList=dao.queryHql ("from Depart where parentId='ff8080815df11660015df237da6b04b5' and name=?0",name);
        return departList.size ()<1 ? name :departList.get (0).getUniquecoding ();
    }
    /**
     * 街道获取
     * @param parentId
     * @param keyword
     * @return
     */
    public List getXhSteer(){
        List<Depart> departList=dao.queryHql ("from Depart where parentId='ff8080815df11660015df237da6b04b5'");
        return departList;
    }

    /**
     * Excel导出工具
     *
     * @param list 查询出需要导出的数据列表
     * @param name 设置数据的列的名称与mapkey位置保持一致
     *             备注：中文名称对应字典中文名称 可以将value转化为字典类型中的中文名称
     * @param mapKey 设置数据的列的list中的key名称与name位置保持一致
     * @param title 导出的execl标题
     * @param args 请求参数
     * @return
     */
    public HSSFWorkbook baseExportExcel(List<HashMap<String,Object>> list, String[] name,String [] mapKey, String title, HashDataPlus args) {
        // 声明一个工作薄
        HSSFWorkbook work = new HSSFWorkbook();
        // 得到excel的第0张表并命名
        Sheet sheet = work.createSheet(title);  //前面所取到的
        sheet.setColumnWidth(0,  (int)((40 + 0.72) * 256));
        sheet.setColumnWidth(5,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(10,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(11,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(12,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(13,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(14,  (int)((11 + 0.72) * 256));
        sheet.setColumnWidth(18,  (int)((20 + 0.72) * 256));
        sheet.setColumnWidth(19,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(20,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(23,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(24,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(25,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(26,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(27,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(28,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(29,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(30,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(31,  (int)((17 + 0.72) * 256));
        sheet.setColumnWidth(32,  (int)((17 + 0.72) * 256)); //1/256
		/*	        sheet.setColumnWidth(1,  (int)((18 + 0.72) * 256));
	        sheet.setColumnWidth(3,  (int)((17 + 0.72) * 256));
	        sheet.setColumnWidth(4,  (int)((17 + 0.72) * 256));

	        sheet.setColumnWidth(6,  (int)((17 + 0.72) * 256));
	        sheet.setColumnWidth(7,  (int)((17 + 0.72) * 256)); */
        // 生成标题样式
        HSSFCellStyle style = work.createCellStyle();
        // 生成栏目样式
        HSSFCellStyle style1 = work.createCellStyle();
        // 生成数据样式
        HSSFCellStyle style2 = work.createCellStyle();
        //样式字体居中
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 水平居中
        style1.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 水平居中
        style1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直居中
        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 水平居中
        //style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        //style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        //style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        //style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        //设置字体
        HSSFFont font = work.createFont();
        font.setFontName("宋体");
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        font.setColor(HSSFColor.RED.index);//HSSFColor.VIOLET.index //字体颜色
        font.setFontHeightInPoints((short)23);


        HSSFFont font1 = work.createFont();
        font1.setFontName("宋体");
        font1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        font1.setFontHeightInPoints((short)12);


        //把字体应用到当前的样式
        style.setFont(font);
        style1.setFont(font1);
        //style1.setFillPattern(HSSFCellStyle.FINE_DOTS );
        style1.setFillPattern(HSSFCellStyle.BIG_SPOTS);
        style1.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        style1.setFillBackgroundColor(HSSFColor.GREY_25_PERCENT.index);
        // style1.setFillBackgroundColor(HSSFColor.GREY_25_PERCENT.index);
        //创建第一行,并合并单元格,写上标题
        Row row = sheet.createRow(0);

        Cell cell = row.createCell(0);
        sheet.addMergedRegion(new CellRangeAddress(0,1,0,name.length));  //first行，last行，first列，last列
        cell.setCellValue(title);
        cell.setCellStyle(style);

        //创建第二行,写上每个分类的名称
        row = sheet.createRow(2);//得到行
        List<Map<String,Object>> getDic=new ArrayList();

        Map<String,String> dicMap=new HashMap<>();
        String namestr="";

        for (int i = 0; i < name.length; i++) {
            //和字典的名称组成
            String dicPath=getDicPath(name[i]);
            if (!"".equals(dicPath)) {
                dicMap.put(mapKey[i], dicPath);
            }
        }


        //将列名数组放入execl中
        for (int i = 0; i < name.length; i++) {
            cell = row.createCell(i);//得到第0个单元格
            sheet.addMergedRegion(new CellRangeAddress(2,3,i,i));
            cell.setCellValue(name[i]);//填充值
            cell.setCellStyle(style1);
        }


        ProgressBar bar = new ProgressBar(args.get(ProgressBar.IDKEY), list.size());
        for (int i = 4; i < list.size() + 4; i++) {
            HashMap<String, Object> map = list.get(i-4);
            row = sheet.createRow(i);//得到行
            for (int j = 0; j < mapKey.length; j++) {
                cell = row.createCell(j);//得到第1个单元格

                //字典值翻译
                //拥有字典的属性名称需要翻译
//                List<HashMap<String,Object>> dicLists=dao.query("select keyword,value from SYS_DICTIONARY where PARENTID=(select id from SYS_DICTIONARY where KEYWORD=?0 and ROWNUM=1)",name[j]);
                if (map.get(mapKey[j])!=null&&dicMap.get(mapKey[j])!=null){
                    String dicCN=dic.getDictionary(dicMap.get(mapKey[j]),1).get(map.get(mapKey[j]))+"";
//                   List<Dictionary> dicList=dao.query(Dictionary.class,"select * from SYS_DICTIONARY where PARENTID=(select id from SYS_DICTIONARY where KEYWORD=?0 and ROWNUM=1) and VALUE=?1",name[j],map.get(mapKey[j]));
                    if (!"null".equals(dicCN)){
                        cell.setCellValue(dicCN);
                    }else {
                        cell.setCellValue(ext.getUnNull("无效的字典参数"+""));
                    }
                }else {
                    cell.setCellValue(ext.getUnNull(map.get(mapKey[j])+""));
                }

                if("birthplace".equals(mapKey[j])&&ext.getUnNull(map.get("birthplace"))!=null&&ext.getUnNull(map.get("birthplace"))!=""){
                    if(map.get("birthplace").toString().length()==12){
                        String birthPlace = pservice.getdivisionName((String)(map.get("birthplace")));
                        cell.setCellValue(birthPlace);
                    }
                }
                //实际居住地址
                if("addresscode".equals(mapKey[j])&&ext.getUnNull(map.get("addresscode"))!=null&&ext.getUnNull(map.get("addresscode"))!=""){
                    String inhabited_area = pservice.getdivisionName((String)(map.get("addresscode")));
                    cell.setCellValue(inhabited_area);
                }
            }
            bar.update((i-3));
        }
        return work;
    }

    /**
     * 获取字典的路径
     * @param keyName 传入字典所在的keyword名
     * @return
     */
    public String getDicPath(String keyName){
        String sql=" select replace(wm_concat(KEYWORD),',','>') name from (select KEYWORD,ROWNUM from SYS_DICTIONARY\n" +
                "   start with KEYWORD = '"+keyName+"' and ROWNUM=1 and VALUE is null" +
                "   connect by prior PARENTID = id order by ROWNUM desc)";
        return daoPlus.selectFeild(sql,"name");
    }
    /**
     * 字典获取
     */
    public ArrayList<HashMap<String,Object>> getDicKV(String parintId){
        return daoPlus.dao ().query ("select keyword key,value from SYS_DICTIONARY where PARENTID=?0 ORDER BY VALUE",parintId);
    }
    /**
     * 字典获取西湖区街道
     */
    public ArrayList<HashMap<String,Object>> getStreetKV(){
        return daoPlus.dao ().query ("SELECT name key,UNIQUECODING VALUE from SYS_DEPART where PARENTID='ff8080815df11660015df237da6b04b5'  or ID='ff8080815d62e0ef015d635df9e2000d' ");
    }


    /**
     * 统一excel导入
     * @param request
     * @return
     */
    
    public Map<String, Object> taxExcel(HttpServletRequest request,Class<?> clz,int colRow,int nameRow) {
        Field field[]=clz.getDeclaredFields();
        HashDataPlus args=new HashDataPlus(request);
        int totalNumber = 0;		//总条数
        int successNumber = 0;		//成功条数
        int errorNumber = 0;		//错误条数
        String errorReason="";		//错误原因
        ProgressBar bar = null;
        boolean result = false;

        HashDataPlus resultdata = new HashDataPlus();
        ListPlus faillist=new ListPlus();
        HashDataPlus fail = new HashDataPlus();
        //先导入一个xls文件，存到附件下
        List<Attachment> attachments = new ArrayList<>();
        User user = ext.getCurUser(request);
//		try {
//			attachments = attachmentService.write(request, "accessory", user.getId(), user.getLoginName());
//		} catch (IllegalStateException e1) {
//			e1.printStackTrace();
//		} catch (IOException e1) {
//			e1.printStackTrace();
//		}
        if (!args.isEqual("fileId","")){
            attachments.add(dao.queryById(Attachment.class,args.getString("fileId")));
        }else{
            ProgressBar.end(request);
            return ext.feedback(false, "未上传数据!");
        }

        ByteArrayInputStream is = null;
        // 同时支持Excel 2003、2007
        try {
            is = new ByteArrayInputStream(attachmentService.getFileBytes(attachments.get(0)));// 文件流
            Workbook workbook = ImportValidation.getWorkbok(is, attachments.get(0).getPath());
            //创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回
            List<String[]> list = new ArrayList<String[]>();
            //设置当前excel中sheet的下标：0开始
            Sheet sheet = workbook.getSheetAt(0);
            //获得当前sheet的开始行
            int firstRowNum  = sheet.getFirstRowNum();
            //获得当前sheet的结束行
            int lastRowNum = sheet.getLastRowNum();

            bar = new ProgressBar(request,lastRowNum);
            String it="";



            for (int i = firstRowNum; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                //如果当前行没有数据，跳出循环
                if(row == null){
                    continue;
                }else{
                    Cell cell = row.getCell(i);
                }

                //获得当前行的开始列
                int firstCellNum = row.getFirstCellNum();
                //获得当前行的列数
                int lastCellNum = sheet.getRow(1).getLastCellNum();

                String[] cells = new String[lastCellNum];
                String[] namecn = new String[lastCellNum];
                String[] colName = new String[lastCellNum];
                for (int j = firstCellNum; j < lastCellNum; j++) {
                    Cell cell = row.getCell(j);
                    if(ImportValidation.getCellValue(cell) == "" || ImportValidation.getCellValue(cell) == null ){
                        continue;
                    }
                    namecn[j]=ImportValidation.getCellValue(row.getCell(j));
                    cells[j] = ImportValidation.getCellValue(cell);
                }
//				for (String cell : cells) {
//					for (Field field1 : field) {
//						if(field1.getName().equals(cell)){
//
//						}
//					}
//				}
                for (int j = firstCellNum; j < lastCellNum; j++) {
                    Cell cell = row.getCell(j);
                    if(ImportValidation.getCellValue(cell) == "" || ImportValidation.getCellValue(cell) == null ){
                        continue;
                    }
                    cells[j] = ImportValidation.getCellValue(cell);
                }
                if(cells[0] !=null && cells[0]!="" ){
                    list.add(cells);
                }
                it+=(i+1)+",";
            }
            String[] line = it.split(",");//获取当前所在行的数组p
            List<Person> person = new ArrayList<Person>();
            SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd");
            IdCardCheck idCardCheck = new IdCardCheck();
            bar = new ProgressBar(request, list.size());
            String[] colNames=list.get(colRow);
            String[] cnName=list.get(nameRow);
            Map<String,String> dicMap=new HashMap<>();
            for (int i=0;i< cnName.length;i++) {
                String dicPath=getDicPath(cnName[i]);
                if (!"".equals(dicPath)) {
                    dicMap.put(colNames[i], dicPath);
                }
            }

            List<Object> listClz=new ArrayList<>();
            for (int m = nameRow+1; m < list.size(); m++) {
                String colName="";
                String values="";
                Object obj=clz.newInstance();
                boolean personExist=false;
                Map<String,Object> objectMap=new HashMap<>();
                for (int i=0;i<list.get(m).length;i++){
                    for (Field field1 : field) {
                        if ("cardid".equals(field1.getName())){
                          Object object=dao.queryHqlRow(" from "+clz.getSimpleName()+" where cardid=?0",list.get(m)[i]);
                            if(object==null){
                                PropertyDescriptor pd = new PropertyDescriptor("id", clz);
                                Method wM = pd.getWriteMethod();//获得写方法
                                wM.invoke(obj, dao.getUUID());//因为知道是int类型的属性，所以传个int过去就是了。。实际情况中需要判断下他的参数类型
                                //判断是否在基础信息中存在
                                personExist=pservice.isExist(list.get(m)[i]);
                            }else{
                                obj=object;
                                errorNumber++;
                                fail.put("","");
                                errorReason="<br><span style='color:white;'>提示：第"+line[m]+"行导入失败,失败原因：该人员信息已存在!</span>";

                                bar.update(m+1);
                            }
                        }

                        if(field1.getName().equals(colNames[i])){
                            PropertyDescriptor pd = new PropertyDescriptor(field1.getName(), clz);
                            Method wM = pd.getWriteMethod();//获得写方法
                            String objValue="";
                            //翻译表格中的参数
                            if (dicMap.get(colNames[i])!=null){
                                objValue=dic.getDictionary(dicMap.get(colNames[i]),0).get(list.get(m)[i])+"";
                            }else {
                                objValue=list.get(m)[i];
                            }
                            wM.invoke(obj,objValue);//因为知道是int类型的属性，所以传个int过去就是了。。实际情况中需要判断下他的参数类型

                            if(!personExist){
                                objectMap.put(colNames[i],objValue);
                            }
                        }
                        if("createdate".equals(field1.getName())){
                            PropertyDescriptor pd = new PropertyDescriptor(field1.getName(), clz);
                            Method wM = pd.getWriteMethod();//获得写方法
                            wM.invoke(obj, new Date());//因为知道是int类型的属性，所以传个int过去就是了。。实际情况中需要判断下他的参数类型
                        }
                        continue;
                    }
                }
                listClz.add(obj);
                if (!personExist){
                    pservice.create(objectMap);
                }
                successNumber++;
                bar.update(m+1);
            }
            totalNumber = list.size();
            if(listClz.size()>=1){
                result =  dao.saveOrUpdateAll(listClz);
            }

            resultdata.put("total", totalNumber);
            resultdata.put("failcount", faillist.size());
            resultdata.put("faildetail", faillist);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            bar.finish();
        }
        Map<String, Object> reback = ext.feedback(result, "总条数："+totalNumber+"条。成功：<span style='color:green;'>"+successNumber+"</span> 条。失败：<span style='color:white;'>"+errorNumber+"</span>条。"+errorReason);
        return reback;
    }


}
