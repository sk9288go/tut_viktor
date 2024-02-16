import os
import json
import compute_rhino3d.Grasshopper as gh
import compute_rhino3d.Util
import rhino3dm
import plotly.graph_objects as go
from visualize_obj import load_obj, save_image_isometric, visualize_obj
from pathlib import Path

compute_rhino3d.Util.url = 'http://localhost:8081/'

workdir = Path(__file__).parent 
print(workdir)
with open(workdir / 'input.json', encoding='utf-8') as f:
    input_params = json.load(f)
    print (input_params)

# Create the input DataTree
input_trees = []
for key, value in input_params.items():
    tree = gh.DataTree(key)
    tree.Append([{0}], [str(value)])
    input_trees.append(tree)

# Evaluate the Grasshopper definition
gh_definition = str(workdir / 'script.gh')
print (gh_definition)
output = gh.EvaluateDefinition(gh_definition, input_trees)

def get_value_from_tree(datatree: dict, param_name: str, index=None):
    """Get first value in datatree that matches given param_name"""
    for val in datatree['values']:
        if val["ParamName"] == param_name:
            try:
                if index is not None:
                    return val['InnerTree']['{0}'][index]['data']
                return [v['data'] for v in val['InnerTree']['{0}']]
            except:
                if index is not None:
                    return val['InnerTree']['{0}'][index]['data']
                return [v['data'] for v in val['InnerTree']['{0;0}']]

# Create a new rhino3dm file and save resulting geometry to file
file = rhino3dm.File3dm()

output_geometry = get_value_from_tree(output, "Geometry", index=0)
output_contextgeometry = get_value_from_tree(output, "ContextGeometry", index=0)
crvs = get_value_from_tree(output, "Curves")

if output_geometry:
    obj = rhino3dm.CommonObject.Decode(json.loads(output_geometry))
    file.Objects.AddMesh(obj)

if output_contextgeometry:
    context_obj = rhino3dm.CommonObject.Decode(json.loads(output_contextgeometry))
    file.Objects.AddMesh(context_obj)

for data in crvs:    
    obj2 = rhino3dm.CommonObject.Decode(json.loads(data))    
    print(obj2)
    file.Objects.Add(obj2)

#라이노 파일을 workdir에 저장 
file_path = str(workdir / 'geometry.obj')  # Path 객체를 문자열로 변환
file.Write(file_path, 7)  # 파일 경로를 문자열로 전달

#output 데이터 합치기 
output_values = {}
for key in ["위치","분석시간","시간스텝","분석 면적","연간 전체 일사량","연간 단위 면적당 일사량","연간 전체 일사량", "일 평균 단위 면적당 일사량"]:
    val = get_value_from_tree(output, key, index=0)
    if val is not None:
        val = val.replace("\"", "")
        output_values[key] = val
        
output_path = workdir / 'output.json'  # Use Path object to construct the path
with open(output_path, 'w', encoding='utf-8') as f:  # Ensure file is opened with UTF-8 encoding
    json.dump(output_values, f, ensure_ascii=False, indent=4)  # Pretty-print with proper encoding

output_path2 = workdir / 'output_save.json'  # Use Path object to construct the path
with open(output_path2, 'w', encoding='utf-8') as f:  # Ensure file is opened with UTF-8 encoding
    json.dump(output_values, f, ensure_ascii=False, indent=4)  # Pretty-print with proper encoding


# 이미지 생성
generated_obj_path =  workdir / 'ColoredOBJ.obj'

if os.path.exists(generated_obj_path):
    vertices, faces, vertex_colors = load_obj(generated_obj_path)
    fig = visualize_obj(vertices, faces, vertex_colors)

    for direction in ['SE', 'SW', 'NE', 'NW']:
        filename = f'isometric_{direction.lower()}.jpg'
        save_image_isometric(fig, direction, filename, workdir, scale_factor=0.5)
