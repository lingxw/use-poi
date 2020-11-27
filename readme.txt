(7)install maven
   apache-maven-3.6.3-bin.zip
   https://blog.csdn.net/u012052268/article/details/78916196
   https://jingyan.baidu.com/article/154b46316f1bfe28ca8f41da.html
   
   settings.xmlを修正(下記の内容をmirrorsの中に追加する)
    <mirror>
     <id>alimaven</id>
     <name>aliyun maven</name>
     <url>http://maven.aliyun.com/nexus/content/groups/public/</url>
     <mirrorOf>central</mirrorOf>
    </mirror>