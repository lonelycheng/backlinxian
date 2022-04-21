-- cw.back_linxian_people definition

CREATE TABLE `back_linxian_people` (
    `name` varchar(100) DEFAULT NULL,
    `id` varchar(100) DEFAULT NULL,
    `tel` varchar(100) DEFAULT NULL,
    `address` varchar(1000) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
    `homeaddr` varchar(100) DEFAULT NULL,
    `backdesc` varchar(1000) DEFAULT NULL,
    `backtime` varchar(100) DEFAULT NULL,
    `xingchengma` varchar(100) DEFAULT NULL,
    `jiankangma` varchar(100) DEFAULT NULL,
    `testresult` varchar(100) DEFAULT NULL,
    `yimiao` varchar(100) DEFAULT NULL,
    `temp` varchar(100) DEFAULT NULL,
    `uploadvillage` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
    `wanggeyuan` varchar(100) DEFAULT NULL,
    `uploaddate` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
    `type` varchar(2) CHARACTER SET utf8 COLLATE utf8_general_ci DEFAULT NULL,
    `method` varchar(100) DEFAULT NULL,
    `person1` varchar(100) DEFAULT NULL,
    `person2` varchar(100) DEFAULT NULL,
    `person3` varchar(100) DEFAULT NULL,
    `person4` varchar(100) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3 COMMENT='疫情返乡人员';