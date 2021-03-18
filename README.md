# EmployeeBPMToAras
![sample](https://github.com/panda0909/EmployeeBPMToAras/blob/master/img/p1.png?raw=true)

從BPM轉出每個人對應的簽核主管

組織層級可由程式修改

```
switch (level_name)
{
    case "董事長級":
        dtRow["cn_level0_manager"] = manager;
        break;
    case "總經理級":
        dtRow["cn_level1_manager"] = manager;
        break;
    case "BU副總經理":
        dtRow["cn_level2_manager"] = manager;
        break;
    case "集團主管":
        dtRow["cn_level3_manager"] = manager;
        break;
    case "處長/協理/副總級":
        dtRow["cn_level4_manager"] = manager;
        break;
    case "資深經理級":
        dtRow["cn_level5_manager"] = manager;
        break;
    case "經理級":
        dtRow["cn_level6_manager"] = manager;
        break;
    case "副理級":
        dtRow["cn_level7_manager"] = manager;
        break;
    case "課長級":
        dtRow["cn_level8_manager"] = manager;
        break;
    case "組長級":
        dtRow["cn_level9_manager"] = manager;
        break;
}
```

1.請輸入BPM連接字串

Data Source=x.x.x.x;Initial Catalog=DatabaseName;User ID=sa;Password=xxxxxxxxxxxxx

2.開始轉檔

3.結束後按下開啟檔案
