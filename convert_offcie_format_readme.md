### 功能说明

1. **多格式支持**：

   - 支持 `.doc → .docx`、`.ppt → .pptx`、`.xls → .xlsx` 三种转换类型
   - 可通过数字菜单选择单个或多个转换类型

2. **交互式菜单**：

   <SHELL>

   ```
   请选择要转换的文件类型（输入数字，多个用逗号分隔）:
   1. Word文档 (.doc → .docx)
   2. PowerPoint演示 (.ppt → .pptx)
   3. Excel表格 (.xls → .xlsx)
   4. 所有类型
   请输入选择（例如 1 或 1,2,3）: 1,3
   ```

3. **增强的错误处理**：

   - 显示详细的转换进度和错误信息
   - 自动跳过已存在的目标文件
   - 捕获 LibreOffice 的错误输出

4. **统计功能**：

   - 最终显示总文件数和成功转换数

   <SHELL>

   ```
   转换完成: 共找到 15 个文件，成功转换 12 个
   ```

5. **符号化输出**：

   - 使用 ✓ ✗ ⚠ 等符号增强可读性
   - 格式化显示转换进度

   

   ### 系统要求

   1. **LibreOffice**：

      - 必须安装 LibreOffice 7.0+
      - 确保 `soffice` 命令可用（Windows需将安装路径加入环境变量）

   2. **Python环境**：

      <SHELL>

      ```
      pip install pathlib
      ```

   3. **跨平台支持**：

      - 已在 Windows/macOS/Linux 测试通过
      - 自动处理不同系统的路径差异