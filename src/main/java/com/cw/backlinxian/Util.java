package com.cw.backlinxian;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class Util {
    public static Map<String, String> jiemaMap = new HashMap<>();
    public static Map<String, String[]> wubaoyiMap = new HashMap<>();
    static {
        jiemaMap.put("前麻峪", "经前麻峪村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。前麻峪村村级调查人员：高建军;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("后麻峪", "经后麻峪村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。后麻峪村村级调查人员：张小;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("赵家石崖", "经赵家石崖村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。赵家石崖村村级调查人员：刘资丰;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("郭家岔", "经郭家岔村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。郭家岔村村级调查人员：高建国;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("城关", "经城关村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。城关村村级调查人员：林德富;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("田家沟", "经田家沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。田家沟村村级调查人员：田晓奇;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("柏树沟", "经柏树沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。柏树沟村村级调查人员：郝子勤;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("东峪沟", "经东峪沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。东峪沟村村级调查人员：刘金峰;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("东峁", "经东峁村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。东峁村村级调查人员：刘志强;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("前甘泉", "经前甘泉村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。前甘泉村村级调查人员：郭泽毅;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("后甘泉", "经后甘泉村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。后甘泉村村级调查人员：冯秋明;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("郭家沟", "经郭家沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。郭家沟村村级调查人员：郭泽锋;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("薛家焉", "经薛家焉村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。薛家焉村村级调查人员：薛全桂;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("都督", "经都督村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。都督村村级调查人员：高兴祚;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("胜利坪", "经胜利坪村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。胜利坪村村级调查人员：高远;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("南塔", "经南塔村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。南塔村村级调查人员：李宏忠;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("槐树塔", "经槐树塔村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。槐树塔村村级调查人员：刘林兆;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("贺家沟", "经贺家沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。贺家沟村村级调查人员：李小军;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("后月镜", "经后月镜村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。后月镜村村级调查人员：刘杰;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("前月镜", "经前月镜村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。前月镜村村级调查人员：刘永军;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("泥沟", "经泥沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。泥沟村村级调查人员：刘金平;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("李家沟", "经李家沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。李家沟村村级调查人员：郭建明;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("万安坪", "经万安坪村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。万安坪村村级调查人员：薛志明;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("上西坡", "经上西坡村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。上西坡村村级调查人员：姜明明;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("黄白塔", "经黄白塔村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。黄白塔村村级调查人员：苗勤喜;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("万安里", "经万安里村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。万安里村村级调查人员：白候平;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("白家沟", "经白家沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。白家沟村村级调查人员：任谈云;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("陈家庄", "经陈家庄村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。陈家庄村村级调查人员：孙利明;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("化林", "经化林村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。化林村村级调查人员：高平山;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("杜家沟", "经杜家沟村村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。杜家沟村村级调查人员：张奴贵;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("凤凰社区", "经凤凰社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。凤凰社区村级调查人员：张彩连;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("盘龙社区", "经盘龙社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。盘龙社区村级调查人员：严小平;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("万安社区", "经万安社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。万安社区村级调查人员：李金梅;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("麻峪苑社区", "经麻峪苑社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。麻峪苑社区村级调查人员：张小;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("从龙社区", "经从龙社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。从龙社区村级调查人员：田兰清;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("太和社区", "经太和社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。太和社区村级调查人员：渠保兰;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("凤城社区", "经凤城社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。凤城社区村级调查人员：赵鹏鸿;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("万安花园社区", "经万安花园社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。万安花园社区村级调查人员：张奴秀;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("晋泰社区", "经晋泰社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。晋泰社区村级调查人员：李改秀;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");
        jiemaMap.put("湫水万安苑社区", "经湫水万安苑社区村调查核实、研判，14天内未出吕梁，本人已签订个人承诺书。显示手机外省漫游回吕梁，是因为与陕西省榆林地市临近，可接收到非吕梁手机基站信号。湫水万安苑社区村级调查人员：刘永红;乡级/单位填报人：成伟;县级健康码解码专班负责人：王凤琪;市级转发人：王自。");


        wubaoyiMap.put("凤凰社区", new String[]{"吕四敏17303581169","薛芳芳13485375111","张彩连 13753389277 ","张彩连 13753389277 "});
        wubaoyiMap.put("胜利坪", new String[]{"吕四敏17303581169","高旭峰13994845589","高远18334770899","高兴祚 13903589027 王进15934016868 刘建梅 15934016000"});
        wubaoyiMap.put("后月镜", new String[]{"吕四敏17303581169","张东荣13513588493","刘杰15135460365","张峰峰 18334803083"});
        wubaoyiMap.put("都督", new String[]{"吕四敏17303581169","李利峰13994819394","李雷雷15935878289","林飞 13994819210 郭林平15935827466 贺文珍 15135462368"});
        wubaoyiMap.put("盘龙社区", new String[]{"薛卫勤13453853939","高旭兰13653588181","严小平 13835853901","严小平 13835853901"});
        wubaoyiMap.put("万安社区", new String[]{"薛卫勤13453853939","高文平13593388499","李金梅 13935830503","李金梅 13935830503"});
        wubaoyiMap.put("东峁", new String[]{"薛卫勤13453853939","刘怀静15135858044","刘志强13934369931","马进军 13593385695 王国强 18303585796孙小林 13485424570"});
        wubaoyiMap.put("杜家沟", new String[]{"薛卫勤13453853939","高建业15835837900杜旭林13935837739","张奴贵13834747336","李金梅 13935830503 李月珍 13753363300李建荣 13935831278"});
        wubaoyiMap.put("化林", new String[]{"薛卫勤13453853939","高平山18634754383","高平山18634754383","王田雨 13513584181 严小平 13835853901 "});
        wubaoyiMap.put("麻峪苑社区", new String[]{"曹华杰15034283435","周艳丽15333589091","张小 13835853435","张小 13835853435"});
        wubaoyiMap.put("柏树沟", new String[]{"曹华杰15034283435","高文平13593388499","郝子勤 13935853330","赵彩虹 15003581861刘林顺 13393580097 樊小旭18234135980"});
        wubaoyiMap.put("前月镜", new String[]{"曹华杰15034283435","马艳强18235898938","刘永军13593361488","郝勇奋 13835852631 苗改平15834394666"});
        wubaoyiMap.put("后麻峪", new String[]{"曹华杰15034283435","郭丽珍15235825225","张小 13835853435","秦利花 13753388786孙旭勤13835805101"});
        wubaoyiMap.put("贺家沟", new String[]{"曹华杰15034283435","高艳芳13934369661 王建亮15234814186","李小军13835814044","张小 13835853435 高丽 13934369797 秦改珍13753363988"});
        wubaoyiMap.put("赵家石崖", new String[]{"曹华杰15034283435","赵瑞云13835832848","刘资丰13503589633","李建峰68988成贵平 13835805369"});
        wubaoyiMap.put("从龙社区", new String[]{"刘志鹏13663684321","曹成仁13935831028","田兰清 18735837730","田兰清 18735837730"});
        wubaoyiMap.put("东峪沟", new String[]{"刘志鹏13663684321","闫全秀15835834688","刘金峰13934369326","高贵平 13503589640 成芳芳15034266490"});
        wubaoyiMap.put("陈家庄", new String[]{"刘志鹏13663684321","齐小伟13835832673","孙利明13593388981","刘守斌 13994817297 "});
        wubaoyiMap.put("后甘泉", new String[]{"刘志鹏13663684321","白雪15735872198","冯秋明13935849378","高顺荣 15235442579 武海艳15035808185"});
        wubaoyiMap.put("郭家沟", new String[]{"刘志鹏13663684321","郭瑜13720930068","郭泽锋18234801394","郭世春 13593388086 薛书亭13313588034"});
        wubaoyiMap.put("薛家焉", new String[]{"刘志鹏13663684321","严翠兵13994819043","薛全桂13753389332","刘金峰 13934369326 李佰选18634755711"});
        wubaoyiMap.put("太和社区", new String[]{"王伟18835805508","高贵芳15235810498","渠保兰13935849514","渠保兰 13935849514"});
        wubaoyiMap.put("南塔", new String[]{"王伟18835805508","贺奇泽13753864449","李宏忠13835832177","白彩虹 15834360066张维清13834753139"});
        wubaoyiMap.put("前麻峪", new String[]{"王伟18835805508","高建新13934369960","高建军15735803222","薛小卫 13994819072 薛艳珍15835883984"});
        wubaoyiMap.put("泥沟", new String[]{"王伟18835805508","刘林森13935892221","刘金平13934369856","高峰 13835818852 李玉新 13934017487"});
        wubaoyiMap.put("郭家岔", new String[]{"王伟18835805508","高 峰13720930064","高建国15635842225","赵锦新 13293665333"});
        wubaoyiMap.put("槐树塔", new String[]{"王伟18835805508","刘泽峰18334835055","刘林兆13994847477","郭忠15135800041"});
        wubaoyiMap.put("凤城社区", new String[]{"张亮亮13935881611","高鹏18334869470","赵鹏鸿15035377536","赵鹏鸿 15035377536"});
        wubaoyiMap.put("万安花园社区", new String[]{"张亮亮13935881611","郭艳艳15035850435","张奴秀15834357637","张奴秀 15834357637"});
        wubaoyiMap.put("田家沟", new String[]{"张亮亮13935881611","田海鹏15386988278","田晓奇15235865333","刘军 13653486999 高改秀15035381005"});
        wubaoyiMap.put("白家沟", new String[]{"张亮亮13935881611","李晓明15834382096 李亚楠13593413772","任谈云13835818810","张向伟6499成云云 15834342860"});
        wubaoyiMap.put("黄白塔", new String[]{"张亮亮13935881611","刘小艳18334887255","张向伟13835804999","赵鹏鸿 15035377536 樊晋丽 13935852506"});
        wubaoyiMap.put("前甘泉", new String[]{"张亮亮13935881611","郭鹏13593385939","郭泽毅13663581117","李芳 15135803021 赵玉斌13835852506"});
        wubaoyiMap.put("万安里", new String[]{"张亮亮13935881611","齐小伟13835832673","李香勤18735819488","马宏 15935185685薛晋平15835834652"});
        wubaoyiMap.put("晋泰社区", new String[]{"曹生勤18735837771","白顺玉13935804546","李改秀18203580082","李改秀 18203580082"});
        wubaoyiMap.put("城关", new String[]{"曹生勤18735837771","王利强13803487072","林德富15135847058","刘毓剑 13935830835 李连连13653487977李有娥15235893758"});
        wubaoyiMap.put("万安坪", new String[]{"曹生勤18735837771","薛新顺15234390094","薛志明13935892260","秦亮亮 13935838732王俊喜62355"});
        wubaoyiMap.put("上西坡", new String[]{"曹生勤18735837771","张唤秀15235883427","姜明明13753389679","刘清泉 13935853166林美平13453116560"});
        wubaoyiMap.put("李家沟", new String[]{"曹生勤18735837771","张晋花15135466316王军军13835800594","郭建明15110338722","刘学民 13453853866王志文19834397673"});
        wubaoyiMap.put("湫水万安苑社区", new String[]{"曹生勤18735837771","杜旭林13935837739","刘永红 15935087000","刘永红 15935087000"});
    }

    public static String getCellValueByCell(Cell cell) {
        //判断是否为null或空串
        if (cell == null || cell.toString().trim().equals("")) {
            return "";
        }
        String cellValue = "";
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC: // 数字
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    cellValue = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 14 || cell.getCellStyle().getDataFormat() == 31 || cell.getCellStyle().getDataFormat() == 57 || cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
                    cellValue = sdf.format(date);
                } else {
//                    double value = cell.getNumericCellValue();
//                    CellStyle style = cell.getCellStyle();
//                    DecimalFormat format = new DecimalFormat();
//                    String temp = style.getDataFormatString();
//                    // 单元格设置成常规
//                    if (temp.equals("General")) {
//                        format.applyPattern("#");
//                    }
//                    cellValue = format.format(value);
                    cell.setCellType(CellType.STRING);
                    cellValue = cell.getStringCellValue();
                }
                break;
            case STRING: // 字符串
                cellValue = cell.getStringCellValue();
                break;
            case BOOLEAN: // Boolean
                cellValue = cell.getBooleanCellValue() + "";
                break;
            case FORMULA: // 公式
                cell.setCellType(CellType.STRING);
                cellValue = cell.getStringCellValue();
                break;
            case BLANK: // 空值
                cellValue = "";
                break;
            case ERROR: // 故障
                cellValue = "ERROR VALUE";
                break;
            default:
                cellValue = "UNKNOWN VALUE";
                break;
        }
        return cellValue.trim();
    }


}
