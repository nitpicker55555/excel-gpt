# excel-gpt
![动画6](https://github.com/nitpicker55555/excel-gpt/assets/91596298/bdfe2700-02ba-4737-8681-496306d329ac)

## How to Use?:
1. Ensure the Python path is already added to the environment variables.
2. Add your own API key to the `openai.key` variable in the `excel.py` file, and set the model type for the program generation (GPT-4 recommended).
3. Ensure `excel.py` and `excel_vba.xlsm` are in the same path.
4. Make Excel trust the file and allow macros to run.
5. Open the `.xlsm` file, click the button in the image to add "prompt" and “Rerun” buttons in the add-ins. After loading, you can delete the buttons on the `.xlsm` spreadsheet.
6. Enter the prompt and click confirm.
7. You can click the Rerun button to rerun the VBA program just generated without consuming tokens.



![image](https://github.com/nitpicker55555/excel-gpt/assets/91596298/de9c044e-e64d-48ee-abf2-28ddbe7010e7)

### Things to Note:
1. The success of the program heavily depends on the quality of the prompt, so please try to clarify the details of the question as much as possible.
2. The program will only pass the prompt to OpenAI; Excel data will not be transferred. Therefore, when writing the prompt, please only write generalized questions, not ones specific to a particular string.
3. Try to operate using this `excel_vba.xlsm` file. Although you can also execute code from GPT in other xlsx files, due to the nature of VBA, the code to generate charts often gets defined to appear on the `.xlsm` spreadsheet.


## 怎么用？：
1. 确保Python路径已经加载到环境变量。
2. 将自己的API key添加到`excel.py`文件中的`openai.key`变量中，并设置需要生成程序的模型型号（推荐GPT-4）。
3. 确保`excel.py`和`excel_vba.xlsm`处于同一路径。
4. 使得Excel信任文件，且允许运行宏函数。
5. 打开`.xlsm`文件，点击图中的按钮，以在加载项增加"prompt"和“Rerun”两个按钮。 加载完后即可删除`.xlsm`表格上的按钮。
6. 输入prompt并点击确认。
7. 可以点击Rerun按钮重新运行刚才生成的VBA程序，且不消耗tokens。

### 需要注意的地方：
1. 程序的成功与否非常依赖于prompt的质量，所以请尽量写清楚问题的细节。
2. 程序只会将prompt传递给OpenAI，Excel的数据并不会被传递，所以在书写prompt时请只书写泛化的问题，而不是针对某个字符串。
3. 尽量使用该`excel_vba.xlsm`表格进行操作，尽管你也可以在其他表格文件中执行来自GPT的代码，但是由于VBA的特性，生成图表的代码经常会被定义出现在`.xlsm`表格上。
