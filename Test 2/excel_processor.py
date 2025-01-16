import pandas as pd
import openai

class ExcelProcessor:
    def __init__(self, file_path):
        """
        Initialize with the Excel file path.
        """
        self.file_path = file_path
        self.data = None

    def read_excel(self):
        """
        Reads the Excel file and loads data.
        """
        try:
            self.data = pd.read_excel(self.file_path)
        except Exception as e:
            raise Exception(f"Error reading Excel file: {e}")

    def process_data(self):
        """
        Processes the data to calculate summaries.
        """
        try:
            average_salary = self.data["Salary"].mean()
            department_distribution = self.data["Department"].value_counts().to_dict()
            return {
                "average_salary": average_salary,
                "department_distribution": department_distribution
            }
        except KeyError as e:
            raise Exception(f"Missing required column: {e}")

    def summarize_with_gpt(self, summary_data):
        """
        Summarizes the processed data using OpenAI's GPT API.
        """
        try:
            openai.api_key = "sk-OOMMKS2qwt1tsKq9BLL7h2hdXNpfY_N-imtNXzWJkqT3BlbkFJ-ZOAM1TlNm1w0zSCOJ2w7ST3p4RSto8FN1X4VufYwA"  # Replace with your API key
            prompt = (
                f"The following data was processed:\n"
                f"Average Salary: {summary_data['average_salary']}\n"
                f"Department Distribution: {summary_data['department_distribution']}\n"
                f"Generate a concise summary."
            )
            response = openai.Completion.create(
                engine="text-davinci-003",
                prompt=prompt,
                max_tokens=100,
                temperature=0.7
            )
            return response["choices"][0]["text"].strip()
        except Exception as e:
            raise Exception(f"Error using OpenAI API: {e}")


if __name__ == "__main__":
    # File path to your Excel file
    file_path = input("/home/dina-iman/Downloads/Employee_Data.xlsx")

    # Create an instance of ExcelProcessor
    processor = ExcelProcessor(file_path)

    try:
        # Read and process the Excel file
        processor.read_excel()
        summary = processor.process_data()

        # Print the processed summary
        print("Processed Summary:")
        print(summary)

        # Generate a GPT-based summary
        gpt_summary = processor.summarize_with_gpt(summary)
        print("\nGPT Summary:")
        print(gpt_summary)
    except Exception as e:
        print(f"Error: {e}")
