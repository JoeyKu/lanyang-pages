import os
import json

def update_index():
    # Path relative to script location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_dir = os.path.join(script_dir, 'default-files')
    
    if not os.path.exists(default_dir):
        print(f"⚠️ Directory {default_dir} does not exist. Creating it...")
        os.makedirs(default_dir)
        
    files = os.listdir(default_dir)
    docx_file = next((f for f in files if f.endswith('.docx') and f != 'index.json'), None)
    pptx_file = next((f for f in files if f.endswith('.pptx')), None)
    
    metadata = {
        "docx": docx_file,
        "pptx": pptx_file
    }
    
    index_path = os.path.join(default_dir, 'index.json')
    with open(index_path, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)
    
    print("✅ index.json updated successfully!")
    print(f"   📄 Word docx: {docx_file or 'None'}")
    print(f"   📊 Slide pptx: {pptx_file or 'None'}")

if __name__ == '__main__':
    update_index()
