import json
import os

class ConfigManager:
    def __init__(self, config_file='settings.json'):
        self.config_file = config_file
        self.config_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(self.config_dir, self.config_file)
        
        self.default_config = {
            'read_order': 'left_to_right',
            'allow_empty': True,
            'filter_keywords': '请示,公告,结果',
            'key_file': '',
            'last_folder': '',
            'last_save_folder': '',  # 新增：上次保存Excel的位置
            'window_position': None,  # 新增：窗口位置
            'window_size': None,      # 新增：窗口大小
            'project_mode': 'same'    # 新增：默认项目处理方式为所有文件夹作为同一项目
        }

    def load_config(self):
        """加载配置，如果出错则返回默认配置"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    saved_config = json.load(f)
                    # 确保所有默认配置项都存在
                    config = self.default_config.copy()
                    config.update(saved_config)
                    # 验证路径是否有效
                    if config['key_file'] and not os.path.exists(config['key_file']):
                        config['key_file'] = ''
                    if config['last_folder'] and not os.path.exists(config['last_folder']):
                        config['last_folder'] = ''
                    return config
            return self.default_config.copy()
        except Exception as e:
            print(f"加载配置出错: {str(e)}")
            return self.default_config.copy()

    def save_config(self, config):
        """保存配置到文件"""
        try:
            # 确保配置目录存在
            os.makedirs(self.config_dir, exist_ok=True)
            
            # 净化配置数据
            clean_config = {
                k: v for k, v in config.items()
                if k in self.default_config
            }
            
            # 保存配置
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(clean_config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存配置失败: {str(e)}")
            return False

    def get_file_dialog_kwargs(self, dialog_type='file'):
        """获取文件对话框的初始参数"""
        config = self.load_config()
        kwargs = {}
        
        if dialog_type == 'file':
            if config['last_folder'] and os.path.exists(config['last_folder']):
                kwargs['initialdir'] = config['last_folder']
        elif dialog_type == 'key_file':
            if config['key_file']:
                initial_dir = os.path.dirname(config['key_file'])
                if os.path.exists(initial_dir):
                    kwargs['initialdir'] = initial_dir
        elif dialog_type == 'save':
            if config['last_save_folder'] and os.path.exists(config['last_save_folder']):
                kwargs['initialdir'] = config['last_save_folder']
        
        return kwargs

    def update_window_state(self, window):
        """更新窗口状态到配置"""
        try:
            config = self.load_config()
            config['window_position'] = window.geometry().split('+')[1:]
            config['window_size'] = window.geometry().split('+')[0]
            self.save_config(config)
        except Exception:
            pass
