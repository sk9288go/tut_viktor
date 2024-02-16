import plotly.graph_objects as go
import numpy as np
import os

# Function to load OBJ file
def load_obj(file_path):
    vertices = []
    faces = []
    vertex_colors = []  # Added to store vertex colors

    with open(file_path, 'r') as obj_file:
        for line in obj_file:
            parts = line.split()
            if len(parts) > 0:
                if parts[0] == 'v':
                    # Vertex data
                    vertices.append([float(parts[1]), float(parts[2]), float(parts[3])])
                    # Check if the vertex has color information (RGB values)
                    if len(parts) >= 6:
                        r, g, b = [int(parts[i]) for i in [4, 5, 6]]
                        vertex_colors.append([r, g, b, 255])  # Set alpha to 255 for RGB
                    else:
                        # If no color information, use white (255, 255, 255, 255)
                        vertex_colors.append([255, 255, 255, 255])
                elif parts[0] == 'f':
                    # Face data
                    face = [int(v.split('/')[0]) for v in parts[1:]]
                    faces.append(face)

    return np.array(vertices), np.array(faces), np.array(vertex_colors)

# Create a Plotly mesh3d figure with vertex colors
def visualize_obj(vertices, faces, vertex_colors):
    fig = go.Figure()

    fig.add_trace(go.Mesh3d(
        x=vertices[:, 0],
        y=vertices[:, 1],
        z=vertices[:, 2],
        i=faces[:, 0] - 1,
        j=faces[:, 1] - 1,
        k=faces[:, 2] - 1,
        vertexcolor=vertex_colors / 255.0,
        flatshading=True,
        lighting=dict(ambient=0.7, diffuse=0.7, specular=0.2)  # Lighting settings
    ))

    # Set the background color and adjust the camera
    fig.update_layout(
        paper_bgcolor="rgba(255,255,255,1)",
        plot_bgcolor='rgba(255,255,255,1)',
        scene=dict(
            xaxis=dict(showgrid=False, backgroundcolor="rgba(255, 255, 255, 1)"),
            yaxis=dict(showgrid=False, backgroundcolor="rgba(255, 255, 255, 1)"),
            zaxis=dict(showgrid=False, backgroundcolor="rgba(255, 255, 255, 1)"),
            xaxis_title='X',
            yaxis_title='Y',
            zaxis_title='Z'
        ),
        scene_camera=dict(eye=dict(x=1.5, y=1.5, z=1.5))
    )

    return fig


def save_image_isometric(fig, direction, filename, save_path, scale_factor=1):
    # Define camera for orthographic projection
    camera = dict(
        projection=dict(type='orthographic'),
        up=dict(x=0, y=0, z=1),
        center=dict(x=0, y=0, z=0)
    )

    # Direction multipliers for different isometric angles
    direction_multipliers = {
        'SE': np.array([1, -1, 1]),
        'SW': np.array([-1, -1, 1]),
        'NE': np.array([1, 1, 1]),
        'NW': np.array([-1, 1, 1]),
    }

    # Apply direction multipliers to the eye position
    eye_multiplier = direction_multipliers.get(direction, np.array([1, -1, 1]))
    eye = eye_multiplier * scale_factor

    # Update the eye position in camera
    camera['eye'] = dict(x=eye[0], y=eye[1], z=eye[2])
    
    # Update figure layout with the camera settings
    fig.update_layout(scene_camera=camera)

    # Use try-except block for error handling
    try:
        fig.write_image(os.path.join(save_path, filename), width=600, height=600)
        print(f"Saved {filename} successfully.")
    except Exception as e:
        print(f"Error saving {filename}: {e}")


