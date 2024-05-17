# spring-boot-starter-easyexcel

<table>
<tr>
  <th>姓名</th>
  <th>年龄</th>
</tr>
  <tr>
    <td>张三</td>
    <td>35</td>
  </tr>
</table>
<pre>

  
    @PostMapping("excel")
    public @Ignored Object excel(@RequestExcel List<Person> people) {
        return people;
    }

    @Data
    public static class Person {

        @ExcelProperty("姓名")
        private String name;

        @ExcelProperty("年龄")
        private String age;
    }
 
</pre>
