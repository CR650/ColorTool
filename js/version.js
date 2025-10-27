/**
 * 版本管理模块
 * 文件说明：管理应用版本信息和版本历史
 * 创建时间：2025-01-09
 * 依赖：无
 */

// 创建全局命名空间
window.App = window.App || {};

/**
 * 版本管理模块
 * 提供版本信息显示和管理功能
 */
window.App.Version = (function() {
    'use strict';

    // 当前版本信息
    const VERSION_INFO = {
        current: '1.7.1',
        releaseDate: '2025-10-27',
        buildNumber: '20251027001'
    };

    // 版本历史记录
    const VERSION_HISTORY = [
        {
            version: '1.7.1',
            date: '2025-10-27',
            changes: [
                '🔧 关键修复：直接映射模式下ColorInfo等工作表的所见即所得问题',
                '✅ 修复文件保存阶段的数据覆盖问题：修改generateUpdatedWorkbook()函数',
                '✅ 总是包含所有UI配置的工作表：ColorInfo、Light、FloodLight、VolumetricFog',
                '✅ 防止原始备份数据对修改值的覆盖：无论Status状态如何都保留用户修改',
                '🎯 UI数据加载逻辑保持不变：根据Status状态决定从哪里读取数据显示',
                '📊 文件保存逻辑修复：总是读取UI上当前显示的值（用户可能已修改）',
                '✨ 完全实现所见即所得：UI上显示什么值，文件里就保存什么值',
                '🔍 添加调试日志：getColorInfoConfigData()和validateRgbValue()中添加日志',
                '🧪 测试验证：直接映射模式下修改ColorInfo参数后正确保存到文件',
                '📝 向后兼容：不影响其他工作表和现有功能的处理逻辑'
            ],
            type: 'patch'
        },
        {
            version: '1.7.0',
            date: '2025-10-14',
            changes: [
                '🔧 重大修复：彻底解决新建主题时所有工作表的跳空行问题',
                '✅ RSC_Theme表修复：Color、Light、ColorInfo、FloodLight、VolumetricFog表智能插入，不再跳空行',
                '✅ UGCTheme表修复：所有UGC工作表（Custom_Ground_Color等）智能插入，不再跳空行',
                '🎯 智能插入算法：自动查找最后有效数据行，新行紧接在有效行之后',
                '🔄 空行利用：如果目标位置有空行，直接替换而不是跳过',
                '📊 模板优化：总是使用有效数据行作为模板，避免复制空行数据',
                '🆕 新建主题逻辑优化：新建主题时强制处理所有工作表，不受Status状态限制',
                '🛡️ 数据保护增强：非目标工作表自动重置为原始数据，防止数据污染',
                '🎨 默认颜色调整：未找到颜色数据时默认设置为蓝色（5C84F1）而非白色',
                '📋 行号一致性：确保所有相关工作表的新主题都在相同行号，便于数据管理'
            ],
            type: 'minor'
        },
        {
            version: '1.6.4',
            date: '2025-09-30',
            changes: [
                '🔧 关键修复：新建主题模式下UI参数显示与保存数据一致性问题',
                '✅ 修复UGC配置显示：新建主题时从源数据正确读取Custom字段值，不再显示默认值',
                '✅ 修复FloodLight IsOn：新建主题时根据源数据FloodLight状态正确勾选IsOn开关',
                '✅ 修复Light/ColorInfo/VolumetricFog配置：新建主题时正确传递isNewTheme参数',
                '🎯 条件读取逻辑优化：所有loadExisting...Config函数支持isNewTheme参数',
                '🔄 FloodLight IsOn智能默认值：当FloodLight状态为1但源数据无IsOn字段时，默认返回1（开启）',
                '🆕 UGCTheme自动处理：新建主题时Custom_Enemy_Color、Custom_Mover_Color、Custom_Mover_Auto_Color总是新建一行',
                '📋 简单复制模式：这3个工作表只复制上一行数据（id+1），不进行复杂字段处理',
                '🛡️ 不影响现有逻辑：仅在新建主题模式下生效，不受映射模式限制',
                '📝 数据一致性保障：确保UI显示的参数值与最终保存到文件的数据完全一致'
            ],
            type: 'patch'
        },
        {
            version: '1.6.3',
            date: '2025-09-25',
            changes: [
                '🆕 新增VolumetricFog配置功能：为RSC_Theme.xls新增VolumetricFog工作表的数据配置界面',
                '🎨 体积雾颜色设置：支持16进制颜色选择器，与FloodLight配置保持一致的交互体验',
                '📐 体积雾尺寸配置：X、Y、Z（宽度、长度、高度）支持0-100范围的整数输入',
                '⚙️ 体积雾效果设置：Density（浓淡）支持0.1精度的小数输入，Rotate（旋转角度）支持-90到90度',
                '🔘 智能开关设计：IsOn使用醒目的toggle switch样式，与FloodLight开关保持一致',
                '🌫️ 橙色主题设计：VolumetricFog配置面板使用独特的橙色渐变主题，与其他配置面板形成视觉区分',
                '🔄 数据转换机制：Density字段输入值自动×10存储，读取时自动÷10显示，确保数据格式正确',
                '📊 完整数据流程：包含getVolumetricFogConfigData、applyVolumetricFogConfigToRow等完整处理函数',
                '🎯 智能验证系统：X/Y/Z（0-100）、Density（0-20）、Rotate（-90到90）范围验证和自动修正',
                '📋 工作表处理集成：将VolumetricFog添加到所有相关函数的targetSheets列表中',
                '🧪 专门测试页面：创建test_volumetricfog_config.html，包含完整的功能验证测试',
                '🔗 工作流程集成：完全集成到现有的主题处理工作流程中，支持新建和更新主题'
            ],
            type: 'minor'
        },
        {
            version: '1.6.2',
            date: '2025-09-25',
            changes: [
                '🔧 关键修复：FloodLight工作表处理集成',
                '📋 修复targetSheets列表：将FloodLight添加到所有相关函数的处理列表中',
                '⚙️ processRSCAdditionalSheets：新建主题时正确处理FloodLight工作表',
                '🔄 updateExistingThemeAdditionalSheets：更新主题时正确处理FloodLight工作表',
                '📁 generateUpdatedWorkbook：生成Excel文件时正确同步FloodLight工作表数据',
                '🧪 专门测试页面：创建test_floodlight_processing.html验证工作表处理集成',
                '📝 更新注释和日志：所有相关函数的注释都包含FloodLight工作表说明',
                '✅ 完整数据流程：确保FloodLight配置在新建和更新主题时都能正确保存到Excel文件'
            ],
            type: 'patch'
        },
        {
            version: '1.6.1',
            date: '2025-09-25',
            changes: [
                '🆕 新增FloodLight配置功能：为RSC_Theme.xls新增FloodLight工作表的数据配置界面',
                '🎨 泛光颜色设置：支持16进制颜色选择器，与Light配置保持一致的交互体验',
                '⚙️ 泛光参数配置：TippingPoint（临界点）和Strength（强度）支持0.1精度的小数输入',
                '🔘 智能开关设计：IsOn和JumpActiveIsLightOn使用醒目的toggle switch样式',
                '🔄 数据转换机制：输入值自动×10存储，读取时自动÷10显示，确保数据格式正确',
                '📊 完整数据流程：包含getFloodLightConfigData、applyFloodLightConfigToRow等完整处理函数',
                '🎯 智能默认值：新建主题时使用表中最后一个主题的FloodLight配置作为默认值',
                '✅ 实时验证系统：TippingPoint（0-5）、Strength（0-10）范围验证和自动修正',
                '🌊 青色主题设计：FloodLight配置面板使用独特的青色渐变主题，与其他配置面板形成视觉区分',
                '🧪 专门测试页面：创建test_floodlight_config.html，包含完整的功能验证测试',
                '🔗 工作流程集成：完全集成到现有的主题处理工作流程中，支持新建和更新主题'
            ],
            type: 'minor'
        },
        {
            version: '1.6.0',
            date: '2025-09-25',
            changes: [
                '🆕 新增AllObstacle.xls文件处理功能：支持全新系列主题的AllObstacle数据自动管理',
                '✨ 智能触发条件：仅在创建全新系列主题时自动处理AllObstacle文件，避免不必要的操作',
                '📊 自动数据填充：在Info工作表中自动添加新记录，包含ID、基础名称、多语言ID、Sort和isFilter字段',
                '🔍 智能字段计算：自动计算新的ID和Sort值（现有最大值+1），提取主题基础名称',
                '🛡️ 完善错误处理：文件不存在、权限失败、重复ID等情况都有适当处理，不影响主流程',
                '💾 格式兼容性：优先保存为XLS格式确保与Java工具兼容，避免Office 2007+ XML格式错误',
                '🔧 多格式保存：支持多种Excel格式尝试和工作簿重建机制，解决XLSX库兼容性问题',
                '📝 详细日志：提供丰富的调试日志输出，便于问题追踪和功能验证',
                '🧪 测试页面：创建专门的AllObstacle功能测试页面，包含完整测试用例和验证清单',
                '📚 文档更新：更新相关文档和JSDoc注释，说明AllObstacle处理的触发条件和字段规则',
                '🔄 向后兼容：完全不影响现有的RSC_Theme和UGCTheme处理逻辑，保持系统稳定性'
            ],
            type: 'major'
        },
        {
            version: '1.5.2',
            date: '2025-09-18',
            changes: [
                '🔧 修复关键问题：RSC_Theme多工作表保存逻辑错误',
                '解决新建主题时Light和ColorInfo工作表数据丢失的问题',
                '修复generateUpdatedWorkbook函数中的条件判断错误',
                '移除工作表名称冲突检查，确保Light和ColorInfo工作表始终被处理',
                '增强工作表处理的调试日志，便于问题诊断',
                '添加Light和ColorInfo配置数据的验证输出',
                '创建多工作表保存问题诊断工具：debug_multi_sheet_save.html',
                '确保UI显示的缓存状态与实际保存的文件内容一致',
                '完善错误处理和数据完整性检查'
            ],
            type: 'patch'
        },
        {
            version: '1.5.1',
            date: '2025-09-18',
            changes: [
                '优化映射模式检测逻辑：调整检测优先级，优先检查"完整配色表"工作表',
                '修复直接映射模式检测问题：解决sourceData对象缺少workbook属性的问题',
                '改进CORS错误处理：增强file://协议访问限制的用户提示和解决方案',
                '完善启动脚本：start_server.bat现在会自动打开浏览器并提供更好的用户体验',
                '新增CORS问题解决方案文档：提供详细的本地访问问题解决指南',
                '优化映射模式检测优先级：完整配色表工作表 > Color工作表 > 默认JSON模式',
                '简化直接映射字段匹配：移除字段数量限制，只要有Color工作表就启用直接映射',
                '增强错误提示：在控制台和UI中提供更清晰的CORS问题解决方案',
                '完善文档：更新README文档，添加CORS问题说明和映射模式详细介绍'
            ],
            type: 'patch'
        },
        {
            version: '1.5.0',
            date: '2025-09-18',
            changes: [
                '新增直接映射模式：支持源数据文件包含Color工作表时的直接字段映射',
                '实现自动映射模式检测：根据源数据文件结构自动选择JSON映射或直接映射模式',
                '添加映射模式指示器：在UI中清晰显示当前使用的映射模式（JSON映射模式/直接映射模式）',
                '扩展颜色通道支持：直接映射模式支持G1-G7和P1-P49的完整范围',
                '简化数据验证：直接映射模式只需检查字段存在性和非空性，无需6位十六进制验证',
                '优化用户体验：源数据文件选择结果中显示检测到的映射模式信息',
                '保持向后兼容：现有JSON映射模式功能完全不受影响，双模式并存',
                '新增核心函数：detectMappingMode()、updateThemeColorsDirect()、findColorValueDirect()',
                '完善文档说明：更新README添加直接映射模式的详细使用指南'
            ],
            type: 'major'
        },
        {
            version: '1.4.8',
            date: '2025-09-18',
            changes: [
                '严格限制RSC_Theme工作表的修改范围，不要改到无关表',
            ],
            type: 'patch'
        },
        {
            version: '1.4.7',
            date: '2025-09-16',
            changes: [
                '修正UGCTheme表Level_show_id字段填值逻辑，根据主题类型采用不同策略',
                '全新主题系列：Level_show_id设置为"新增主题ID - 1"',
                '现有主题系列：Level_show_id继续使用"上一行数据值 + 1"的递增逻辑',
                '基于智能多语言配置系统实现主题类型自动识别',
                '添加详细的填值过程日志输出，包含计算来源和处理步骤',
                '确保与多语言配置功能完全兼容，不影响现有功能',
                '提升数据处理准确性，避免全新主题错误延续递增逻辑',
                '完善错误处理机制，处理列不存在和数据异常情况'
            ],
            type: 'patch'
        },
        {
            version: '1.4.6',
            date: '2025-09-12',
            changes: [
                '新增源数据文件选择结果显示功能，提供文件名、大小、时间和数据行数的详细信息',
                '优化UI界面，移除冗余的文件状态栏和右侧信息区域，界面更加简洁',
                '修改最终操作指引弹窗，使用固定路径显示，提供更明确的操作指引',
                '完善文件选择状态管理，在错误情况下自动隐藏结果显示',
                '集成文件选择结果到重置功能，确保界面状态一致性',
                '使用绿色主题的成功状态显示，与项目整体UI风格保持一致',
                '添加实时文件信息更新，包括格式化的文件大小和精确的选择时间',
                '提供完整的错误处理机制，确保用户体验的流畅性'
            ],
            type: 'minor'
        },
        {
            version: '1.4.5',
            date: '2025-09-12',
            changes: [
                '更新图案'
            ],
            type: 'minor'
        },
        {
            version: '1.4.4',
            date: '2025-09-10',
            changes: [
                '新增BallSpec钻石高光颜色配置功能，支持钻石高光颜色的RGB设置',
                '在ColorInfo配置面板中添加BallSpecR、BallSpecG、BallSpecB三个新字段',
                '完整实现BallSpec字段的UI界面、数据收集、数据应用、验证和默认值处理',
                '集成BallSpec字段到getColorInfoConfigData()数据收集函数',
                '集成BallSpec字段到applyColorInfoConfigToRow()数据应用函数',
                '集成BallSpec字段到initColorInfoValidation()验证函数',
                '集成BallSpec字段到getLastThemeColorInfoConfig()默认值函数',
                '集成BallSpec字段到updateRgbColorPreview()颜色预览函数',
                '修复loadExistingColorInfoConfig()函数缺少BallSpec字段映射的问题',
                '确保新建和更新主题时BallSpec配置都能正确保存到RSC_Theme文件',
                '提供完整的BallSpec颜色选择器，支持可视化颜色选择和RGB值自动拆分',
                '添加BallSpec字段的实时验证，确保RGB值在0-255范围内',
                '完善ColorInfo配置功能，现在支持钻石颜色、反光颜色、高光颜色和雾效设置'
            ],
            type: 'minor'
        },
        {
            version: '1.4.3',
            date: '2025-09-10',
            changes: [
                '修复Light和ColorInfo数据保存的关键问题，确保配置数据正确写入Excel文件',
                '解决新建主题时Light和ColorInfo工作表不创建新行的问题',
                '修复更新现有主题时Light和ColorInfo配置被忽略的问题',
                '修复generateUpdatedWorkbook函数只处理主工作表的缺陷，现在正确处理所有工作表',
                '修复processRSCAdditionalSheets函数在新建主题时没有更新rscAllSheetsData的问题',
                '新增updateExistingThemeAdditionalSheets函数，专门处理现有主题的Light和ColorInfo更新',
                '完善数据流程，确保无论新建还是更新主题，Light和ColorInfo配置都能正确保存',
                '移除过时的功能限制说明，现在所有配置功能都完全可用',
                '优化数据同步逻辑，确保workbook.Sheets和rscAllSheetsData保持一致',
                '添加详细的调试日志，便于问题诊断和数据流程追踪'
            ],
            type: 'major'
        },
        {
            version: '1.4.2',
            date: '2025-09-10',
            changes: [
                '初步最终集成工具完成',
            ],
            type: 'minor'
        },
        {
            version: '1.4.1',
            date: '2025-01-10',
            changes: [
                '新增RSC_Theme ColorInfo配置功能，支持钻石颜色和远景雾颜色设置',
                '实现RGB颜色字段的可视化配置，支持钻石颜色、反光颜色、远景雾颜色',
                '添加雾效距离参数配置，支持FogStart和FogEnd的范围验证',
                '为跳板边框图案ID添加可视化选择器，与地板边框保持一致的交互体验',
                '修复玻璃透明度默认值问题，确保新建主题时默认值为50',
                '优化Light和ColorInfo配置的默认值逻辑，新建主题时使用表中最后一个主题的数据',
                '移除跳板边框图案的0/1限制注释，支持任意有效的边框图案ID选择',
                '完善RGB颜色选择器，支持颜色自动拆分为R、G、B三个数值字段',
                '增强配置面板的用户体验，提供实时颜色预览和验证反馈'
            ],
            type: 'minor'
        },
        {
            version: '1.4.0',
            date: '2025-01-09',
            changes: [
                '新增RSC_Theme Light配置功能，支持完整的光照参数设置',
                '实现明度偏移、高光等级、光泽度和高光颜色的可视化配置',
                '添加专业的16进制颜色选择器，支持渐变色板和预设颜色',
                '集成实时数值验证系统，自动检查范围限制并提供错误提示',
                '修复UGC图案选择器的事件绑定问题，确保按钮响应正常',
                '优化Light配置面板UI设计，提供清晰的字段说明和使用提示',
                '添加Light配置的数据加载和保存功能，支持现有主题值显示'
            ],
            type: 'major'
        },
        {
            version: '1.3.1',
            date: '2025-09-09',
            changes: [
                '更换主题Unity工具链接',
                '更新导航栏UI'
            ],
            type: 'minor'
        },
        {
            version: '1.3.0',
            date: '2025-09-09',
            changes: [
                '增加导航站',
            ],
            type: 'minor'
        },
        {
            version: '1.2.9',
            date: '2025-09-09',
            changes: [
                '更换静态部署工具为GitLab',
            ],
            type: 'minor'
        },
        {
            version: '1.2.8',
            date: '2025-09-08',
            changes: [
                '手动改表的操作写到工具的自动化流程中',
            ],
            type: 'minor'
        },
        {
            version: '1.2.7',
            date: '2025-09-08',
            changes: [
                '实现多语言配置智能检测逻辑，支持主题名称相似性检测',
                '添加同系列主题自动复用多语言配置功能',
                '优化多语言配置面板显示逻辑，仅在全新主题系列时显示',
                '新增主题类型指示器，提供清晰的状态反馈',
                '支持中英文主题名称的智能基础名称提取',
                '完善UGCTheme处理中的多语言ID复用机制'
            ],
            type: 'minor'
        },
        {
            version: '1.2.6',
            date: '2025-09-08',
            changes: [
                '实现UGCTheme文件多语言ID自动填充功能',
                '在processUGCTheme函数中添加LevelName列的智能处理',
                '支持多种LevelName列名格式的自动识别',
                '用户输入有效时使用多语言ID，无效时使用上一行数据',
                '添加详细的日志输出和错误处理机制',
                '修复多语言配置UI样式问题，提高文字可见性'
            ],
            type: 'minor'
        },
        {
            version: '1.2.5',
            date: '2025-09-08',
            changes: [
                '新增多语言配置UI功能，仅在创建新主题时显示',
                '添加主题显示名称输入框和多语言ID输入框',
                '集成在线多语言表链接，提供清晰的操作指引',
                '实现多语言配置验证逻辑，不影响现有核心功能',
                '优化用户界面，保持与现有风格一致'
            ],
            type: 'minor'
        },
        {
            version: '1.2.4',
            date: '2025-09-08',
            changes: [
                'UGCTheme表的Level_show_id与Level_show_bg_ID列不再需要手动操作,而是放到当前工具内的自动化流程里面了',
            ],
            type: 'minor'
        },
        {
            version: '1.2.3',
            date: '2025-09-05',
            changes: [
                '重构映射逻辑：优先检查颜色代码，再验证RC通道有效性',
                '改进映射处理：支持更灵活的颜色代码到RC通道的映射',
                '优化跳过逻辑：只有颜色代码为空时才跳过，RC通道无效时使用默认值',
                '增强日志输出：提供更详细的映射处理统计信息',
                '提升映射覆盖率：处理更多映射表中的有效颜色代码'
            ],
            type: 'minor'
        },
        {
            version: '1.2.1',
            date: '2025-09-05',
            changes: [
                '文字修改'
            ],
            type: 'minor'
        },
        {
            version: '1.2.0',
            date: '2025-09-05',
            changes: [
                '修复主题提取逻辑，正确从第6行开始读取主题数据',
                '优化更新模式下UGCTheme文件处理逻辑',
                '改进错误提示，明确指出缺少的具体文件',
                '完善主题数据处理流程和用户反馈',
                '添加功能限制说明：部分数据如雾色花纹使用上一个主题的数据'
            ],
            type: 'minor'
        },
        {
            version: '1.1.1',
            date: '2025-09-05',
            changes: [
                '修复Excel文件格式兼容性问题',
                '统一使用.xls格式以兼容Unity工具',
                '解决Apache POI HSSF库读取错误',
                '添加版本显示功能',
                '优化文件保存和下载流程'
            ],
            type: 'major'
        },
        {
            version: '1.0.0',
            date: '2025-01-08',
            changes: [
                '初始版本发布',
                '实现颜色主题数据处理功能',
                '支持RSC_Theme和UGCTheme文件管理',
                '添加File System Access API支持',
                '实现数据预览和Sheet选择功能',
                '添加响应式界面设计',
                '支持拖拽文件上传',
                '实现数据同步和映射功能'
            ],
            type: 'major'
        }
    ];

    /**
     * 初始化版本模块
     */
    function init() {
        displayVersion();
        console.log(`颜色主题管理工具 v${VERSION_INFO.current} (${VERSION_INFO.releaseDate})`);
    }

    /**
     * 显示版本信息
     */
    function displayVersion() {
        const versionElement = document.getElementById('versionNumber');
        if (versionElement) {
            versionElement.textContent = `v${VERSION_INFO.current}`;
            versionElement.title = `版本 ${VERSION_INFO.current}\n发布日期: ${VERSION_INFO.releaseDate}\n构建号: ${VERSION_INFO.buildNumber}`;
        }
    }

    /**
     * 获取当前版本信息
     * @returns {Object} 版本信息对象
     */
    function getCurrentVersion() {
        return { ...VERSION_INFO };
    }

    /**
     * 获取版本历史
     * @returns {Array} 版本历史数组
     */
    function getVersionHistory() {
        return [...VERSION_HISTORY];
    }

    /**
     * 获取最新的更新内容
     * @returns {Object} 最新版本的更新内容
     */
    function getLatestChanges() {
        return VERSION_HISTORY[0] || null;
    }

    /**
     * 检查是否为新版本
     * @param {string} lastKnownVersion - 上次已知版本
     * @returns {boolean} 是否为新版本
     */
    function isNewVersion(lastKnownVersion) {
        if (!lastKnownVersion) return true;
        
        const current = VERSION_INFO.current.split('.').map(Number);
        const last = lastKnownVersion.split('.').map(Number);
        
        for (let i = 0; i < Math.max(current.length, last.length); i++) {
            const currentPart = current[i] || 0;
            const lastPart = last[i] || 0;
            
            if (currentPart > lastPart) return true;
            if (currentPart < lastPart) return false;
        }
        
        return false;
    }

    /**
     * 显示版本更新通知
     */
    function showUpdateNotification() {
        const lastVersion = localStorage.getItem('colorTool_lastVersion');
        
        if (isNewVersion(lastVersion)) {
            const changes = getLatestChanges();
            if (changes && App.Utils) {
                const changesList = changes.changes.slice(0, 3).join('、');
                App.Utils.showStatus(
                    `🎉 已更新到 v${VERSION_INFO.current}！主要更新：${changesList}`, 
                    'success', 
                    5000
                );
            }
            
            // 保存当前版本
            localStorage.setItem('colorTool_lastVersion', VERSION_INFO.current);
        }
    }

    // 暴露公共接口
    return {
        init: init,
        getCurrentVersion: getCurrentVersion,
        getVersionHistory: getVersionHistory,
        getLatestChanges: getLatestChanges,
        isNewVersion: isNewVersion,
        showUpdateNotification: showUpdateNotification
    };

})();

// 模块加载完成日志
console.log('Version模块已加载 - 版本管理功能已就绪');
