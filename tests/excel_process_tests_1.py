import pytest
from pytest_mock import mocker
import sys
sys.path.append('../')
from src.excel_process_2 import ExcelPreProcessor
import base64

@pytest.fixture
def processor():
    excel_test_file = './large_test_1/large_test_1.xlsx'
    csv_test_file = './large_test_1/large_test_1.csv'
    txt_test_file = './large_test_1/large_test_1.txt'
    image_test_path = './large_test_1/large_test_1.png'
    processor = ExcelPreProcessor(excel_test_file, csv_test_file, txt_test_file, image_test_path)
    return processor

def test_init(processor):
    """Test the initialization of the ExcelPreProcessor."""
    assert processor.excel_filename == './large_test_1/large_test_1.xlsx'
    assert processor.csv_filename == './large_test_1/large_test_1.csv'
    assert processor.txt_filename == './large_test_1/large_test_1.txt'
    assert processor.image_path == './large_test_1/large_test_1.png'

def test_process_excel_to_csv(processor, mocker):
    """Test converting Excel to CSV."""
    mock_wb = mocker.patch('openpyxl.load_workbook')
    mock_ws = mocker.MagicMock()
    mock_wb.return_value.active = mock_ws
    mock_ws.iter_rows.return_value = [[1, 2, 3], [4, 5, 6]]
    
    mock_csv_open = mocker.patch("builtins.open", mocker.mock_open())
    mock_csv_writer = mocker.patch("csv.writer")
    mock_csv_writer_instance = mocker.MagicMock()
    mock_csv_writer.return_value = mock_csv_writer_instance

    processor.process_excel_to_csv()

    mock_wb.assert_called_once_with('./large_test_1/large_test_1.xlsx', data_only=True)
    mock_ws.iter_rows.assert_called_once_with(values_only=True)
    mock_csv_open.assert_called_once_with('./large_test_1/large_test_1.csv', 'w', newline="")
    mock_csv_writer_instance.writerow.assert_any_call([1, 2, 3])
    mock_csv_writer_instance.writerow.assert_any_call([4, 5, 6])

def test_write_csv_to_txt(processor, mocker):
    """Test writing CSV content to TXT file."""
    mock_open_txt = mocker.patch("builtins.open", mocker.mock_open())
    mock_open_csv = mocker.patch("builtins.open", mocker.mock_open(read_data="1,2,3\n4,5,6"))

    processor.write_csv_to_txt()

    mock_open_txt.assert_called_once_with('./large_test_1/large_test_1.txt', 'w')
    mock_open_csv.assert_called_once_with('./large_test_1/large_test_1.csv')
    
    mock_open_txt().write.assert_any_call("1 2 3\n")
    mock_open_txt().write.assert_any_call("4 5 6\n")

def test_encode_image(processor, mocker):
    """Test encoding an image to base64."""
    mock_open_image = mocker.patch("builtins.open", mocker.mock_open(read_data=b"test_image_data"))
    encoded_image = processor.encode_image()
    assert encoded_image == base64.b64encode(b"test_image_data").decode('utf-8')
    mock_open_image.assert_called_once_with('./large_test_1/large_test_1.png', 'rb')

def test_run(processor, mocker):
    """Test the complete run method."""
    mock_process_excel_to_csv = mocker.patch.object(processor, 'process_excel_to_csv')
    mock_write_csv_to_txt = mocker.patch.object(processor, 'write_csv_to_txt')
    mock_query_openai = mocker.patch.object(processor, 'query_openai')
    mock_system_prompt = mocker.patch.object(processor, 'system_prompt', return_value="system_prompt")
    mock_user_prompt = mocker.patch.object(processor, 'user_prompt', return_value="user_prompt")

    processor.run()

    mock_process_excel_to_csv.assert_called_once()
    mock_write_csv_to_txt.assert_called_once()
    mock_query_openai.assert_called_once_with("system_prompt", "user_prompt")

def test_system_prompt(processor):
    """Test the system prompt generation."""
    prompt = processor.system_prompt()
    assert "The following image contains tables in a spreadsheet." in prompt
    assert "provide me the data in each table in JSON format." in prompt

def test_user_prompt(processor, mocker):
    """Test the user prompt generation."""
    mocker.patch.object(processor, 'txt_filename', 'test_1.txt')
    prompt = processor.user_prompt('test_1.txt')
    assert "Consider the following txt representation the spreadsheet in csv format before parsing the image:" in prompt
    assert "Image is provided in base64 format:" in prompt

def test_query_openai(processor, mocker):
    """Test the OpenAI query method."""
    mock_encode_image = mocker.patch.object(processor, 'encode_image', return_value="encoded_image_data")
    mock_openai_create = mocker.patch.object(processor.client.chat.completions, 'create')
    mock_response = mocker.MagicMock()
    mock_response.choices[0].message.content = "response_content"
    mock_openai_create.return_value = mock_response

    processor.query_openai("system_prompt", "user_prompt")

    mock_openai_create.assert_called_once_with(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": [{"type": "text", "text": "system_prompt"}]},
            {"role": "user", "content": [
                {"type": "text", "text": "user_prompt"},
                {"type": "image_url", "image_url": {"url": "data:image/jpeg;base64,encoded_image_data"}}
            ]}
        ],
        temperature=0
    )

if __name__ == "__main__":
    pytest.main()
