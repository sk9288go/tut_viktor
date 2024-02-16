import json
from datetime import date
from pathlib import Path
from io import BytesIO
from viktor.utils import convert_word_to_pdf, render_jinja_template
from viktor.external.word import WordFileTag, WordFileImage, render_word_file
from viktor import ViktorController, File
from viktor.parametrization import ViktorParametrization, Step, Section, Text, NumberField, LineBreak, \
     OptionField, FileField, TextField, DateField
from viktor.views import WebView, WebResult, GeometryAndDataView, GeometryAndDataResult, \
   PDFView, PDFResult, DataGroup, DataItem
from viktor.external.generic import GenericAnalysis
#from viktor.core import Storage



class Parametrization(ViktorParametrization):

    #whats next parameters
    step_1 = Step("Intro", views=["whats_next"])    
    
    step_2 = Step('분석요소 설정', views =["run_grasshopper"])    
    step_2.section_1 = Section('파일 업로드')
    step_2.section_1.input_1 = OptionField('EPW 기상데이터 선택', options=['서울','인천','강릉','울산','부산','광주','제주','서귀포','대전','보령','울진','대구','창원','춘천','원주','전주','청주','목포'], default='서울')
    step_2.section_1.input_2 = FileField('분석  obj 업로드', file_types=['.obj', '.3dm'], max_size=50_000)
    step_2.section_1.input_3 = FileField('주변 건물 obj 업로드', file_types=['.obj', '.3dm'], max_size=50_000)

    step_2.section_2 = Section ('분석 시간')
    step_2.section_2.input_1 = Text("분석 시작과 마감 월, 일 ,시를 지정해주세요\n"
                                   "\n(기본값 1월1일6시 ~ 12월31일20시)") 
    step_2.section_2.input_2 = NumberField("분석 시작 월",default=1, min=1, max=12, description="분석 마감 월")
    step_2.section_2.input_3 = NumberField("분석 시작 일",default=1, min=1, max=31, description="분석 시작 일")
    step_2.section_2.input_4 = NumberField("분석 시작 시",default=6, min=1, max=24, description="분석 시작 시")
    step_2.section_2.input_5 = NumberField("분석 마감 월",default=12, min=1, max=12, description="분석 마감 월")
    step_2.section_2.input_6 = NumberField("분석 마감 일",default=31, min=1, max=31, description="분석 마감 일")
    step_2.section_2.input_7 = NumberField("분석 마감 시",default=20, min=1, max=24, description="분석 마감 시")  
    step_2.section_2.input_8 = LineBreak()
    step_2.section_2.input_9 = OptionField("타임스텝",options=[1,2,3,4,5,6,10,12,15,20,30,60],default=1, description="""1시간을 얼마나 나누어 계산할지를 지정 \\\n (적을수록 계산시간이 빨라짐)""")
    
    step_2.section_3 = Section ('분석 세팅값')
    step_2.section_3.input_1 = NumberField("그리드 크기",default=10, min=1, max=50, description="""분석 면을 나눌 그리드 사이즈\\\n (적을수록 계산시간이 빨라짐)""")
    step_2.section_3.input_2 = NumberField("오프셋 거리",default=1, min=1, max=200, description="""분석 면에서 바깥쪽으로 계산점을 이동하는거리 \\\n(통상적으로 적은 양수의 거리로 주변건물의 영향이 없어야 함)""") 
    step_2.section_3.input_3 = LineBreak()
    step_2.section_3.input_4 = NumberField("범례 하한값",default=1, min=1, max=200, description="""범례의 하한값을 지정 \\\n (범례에 해당 최소값이 반영됩니다)""")
    step_2.section_3.input_5 = NumberField("범례 최대값",default=1200, min=1, max=1500,description="""범례의 상한값을 지정 \\\n (범례에 해당 최대값이 반영됩니다)""")
    step_2.section_3.input_6 = NumberField("color_index",default=3, min=1, max=26)

    step_3 = Step('분석레포트', views =["pdf_view"])   

    step_3.intro = Text("# Invoice app 💰 \n This app makes an invoice based on your own Word template")

    step_3.client_name = TextField("Name of client")
    step_3.company = TextField("Company name")
    step_3.lb1 = LineBreak()  # This is just to separate fields in the parametrization UI
    step_3.date = DateField("Date")
    

class Controller(ViktorController):
    label = 'My Entity Type'
    parametrization = Parametrization(width=30)
    
    @WebView("Intro", duration_guess=1)
    def whats_next(self, **kwargs):
        html_path = Path(__file__).parent / "info_page" / "html_template.html"        
        # HTML 템플릿 파일 읽기
        with open(html_path, 'r', encoding='utf-8') as template_file:
            template_content = template_file.read()
        #html_file = render_jinja_template(template_content)        
        # WebResult로 HTML 내용 반환
        return WebResult(html=template_content)
    
    @GeometryAndDataView("Geometry", duration_guess=30, update_label='run_grasshopper')    
    def run_grasshopper(self, params, **kwargs):
        #storage = Storage()
        json_path = self.create_json_input(params)
        files = [('model.obj', BytesIO(params.step_2.section_1.input_2.file.getvalue_binary())),
                 ('context.obj', BytesIO(params.step_2.section_1.input_3.file.getvalue_binary()))]

        generic_analysis = GenericAnalysis(files=files, executable_key="run_grasshopper", output_filenames=["geometry.obj", "output.json"])
        generic_analysis.execute(timeout=2600)

        rhino_3dm_file = generic_analysis.get_output_file("geometry.obj", as_file=True)
        output_values: File = generic_analysis.get_output_file("output.json", as_file=True)                 
        output_dict = json.loads(output_values.getvalue())            
        
        data_group = DataGroup(
            *[DataItem(key.replace("_", " "), val) for key, val in output_dict.items()]
        )
        return GeometryAndDataResult(geometry=rhino_3dm_file, geometry_type="obj", data=data_group)
          
    def generate_word_document(self, params):
            # Create emtpy components list to be filled later
        #torage = Storage()
        components = []
            # Fill components list with data
        components.append(WordFileTag("Client_name", str("김철수")))
        components.append(WordFileTag("company", params.step_2.section_1.input_1))
        components.append(WordFileTag("Date", str(params.step_3.date))) # Convert date to string format
        
                
        
        json_path = Path(__file__).parent / "files" / "output_save.json"

        analysis_text = ""
        with open(json_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            for key, value in data.items():
                analysis_text += f"{key}: {value}\n"  # Aggregate data into one string
        components.append(WordFileTag("Analysis", analysis_text.strip()))  # Remove last newline
        
        image_path_ne = Path(__file__).parent / "files" / "isometric_ne.jpg"
        image_path_nw = Path(__file__).parent / "files" / "isometric_nw.jpg"
        image_path_se = Path(__file__).parent / "files" / "isometric_se.jpg"
        image_path_sw = Path(__file__).parent / "files" / "isometric_sw.jpg"

        # Append NE image
        with open(image_path_ne, 'rb') as img_file:
            word_file_figure = WordFileImage(img_file, "figure_ne", height=150)
            components.append(word_file_figure)

        # Append NW image
        with open(image_path_nw, 'rb') as img_file:
            word_file_figure = WordFileImage(img_file, "figure_nw", height=150)
            components.append(word_file_figure)

        # Append SE image
        with open(image_path_se, 'rb') as img_file:
            word_file_figure = WordFileImage(img_file, "figure_se", height=150)
            components.append(word_file_figure)

        # Append SW image
        with open(image_path_sw, 'rb') as img_file:
            word_file_figure = WordFileImage(img_file, "figure_sw", height=150)
            components.append(word_file_figure)


            # Get path to template and render word file
        template_path = Path(__file__).parent / "files" / "Template.docx"
        with open(template_path, 'rb') as template:
             word_file = render_word_file(template, components)

        return word_file
    
    @PDFView("PDF viewer", duration_guess=5)
    
    def pdf_view(self, params, **kwargs):
        word_file = self.generate_word_document(params)

        with word_file.open_binary() as f1:
            pdf_file = convert_word_to_pdf(f1)

        return PDFResult(file=pdf_file)

    def create_json_input(self, params):
        # 파라미터를 JSON 형식으로 변환
        json_data = {
        "filePath": "C:/skp/viktor-apps/tut_viktor/files/model.obj",
        "contextPath": "C:/skp/viktor-apps/tut_viktor/files/context.obj",
        "save_Path": "C:/skp/viktor-apps/tut_viktor/files/",
        "epw_loc": params.step_2.section_1.input_1,
        "start_month": params.step_2.section_2.input_2,
        "start_day": params.step_2.section_2.input_3,
        "start_hour": params.step_2.section_2.input_4,
        "end_month": params.step_2.section_2.input_5,
        "end_day": params.step_2.section_2.input_6,
        "end_hour": params.step_2.section_2.input_7,
        "timestep": params.step_2.section_2.input_9,
        "grid_size": params.step_2.section_3.input_1,
        "offset_dist": params.step_2.section_3.input_2,
        "analysis_min": params.step_2.section_3.input_4,
        "analysis_max": params.step_2.section_3.input_5,
        "color_index": params.step_2.section_3.input_6 #이 값은 예시입니다. 실제 애플리케이션에 따라 조정해야 합니다.
    }
        json_path = Path(__file__).parent / "files/input.json"
        with json_path.open("w", encoding='utf-8') as json_file:
            json.dump(json_data, json_file, ensure_ascii=False, indent=4)

        return json_path

