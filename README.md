# excel-gpt
![动画6](https://github.com/nitpicker55555/excel-gpt/assets/91596298/bdfe2700-02ba-4737-8681-496306d329ac)

# How to Use?:

1. **Ensure Python path is already added to the environment variables.**
2. **Add your API key to the openai.key variable in the excel.py file, and set the model type required for generating the program (gpt-4 recommended).**
3. **Ensure that excel.py and excel_vba.xlsm are in the same path.**
4. **Make Excel trust the file and allow macro functions to run.**
5. **Open the xlsm file and click the button in the figure to add "prompt" and "Rerun" buttons to the Add-ins.**
6. **Enter the prompt and click confirm.**
7. **You can click the Rerun button to rerun the VBA program just generated without consuming tokens.**

## Points to Note:

1. **The success of the program heavily depends on the quality of the prompt, so please try to clearly describe the details of the issue.**
2. **The program will only pass the prompt to OpenAI; Excel data will not be transferred, so when writing prompts, please only write about generalized issues, not specific to a particular string.**
