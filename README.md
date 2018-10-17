# Huawei_FW_E1000-9000_USG6600-9500-Cli_ScriptGeneration
HUAWEI firewall script generator（SNAT,DNAT,Security-policy）


防火墙配置脚本生成工具：
  可以根据在service_IpPort_Open statistics中输入的外网映射地址、外部端口、私网地址、内网端口、负载均衡地址、前置端口，在ConfigurationTemplate_V2.10
  中生成防火墙配置文本，文本可直接通过CLI copy导入设备。
  
PoolCopy_V2.0-mod.py实际起到exc表格数据处理、copy的功能。


注意事项：

    1.PoolCopy_V2.0-mod.py文件依赖于service_IpPort_Open statistics.xlsx，ConfigurationTemplate_V2.10.xlsx文件。
    
    2.service_IpPort_Open statistics.xlsx用于存储和输入ip、port、name等信息，ConfigurationTemplate_V2.10.xlsx为配置生成模板。
    
    3.初次使用注意修改PoolCopy_V2.0-mod.py文件内service_IpPort_Open statistics.xlsx，ConfigurationTemplate_V2.10.xlsx的路径位置，推荐使用绝对路径。
    
    
    
    It is not pretty (sorry!) but it does the job for me.
