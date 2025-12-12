// Excel导入导出功能、模态框交互

// DOM元素
const addContactBtn = document.getElementById('addContactBtn');
const contactModal = document.getElementById('contactModal');
const confirmModal = document.getElementById('confirmModal');
const contactForm = document.getElementById('contactForm');
const exportBtn = document.getElementById('exportBtn');
const importFile = document.getElementById('importFile');
const addMethodBtn = document.getElementById('addMethodBtn');
const contactMethodsContainer = document.getElementById('contactMethods');

// 共享变量
let contactToDelete = null;
let currentEditId = null;

// 设置事件监听器 - B同学负责的部分
function setupEventListeners_B() {
    // 添加联系人按钮
    addContactBtn.addEventListener('click', () => openContactModal());

    // 导出功能
    exportBtn.addEventListener('click', exportToExcel);

    // 导入功能
    importFile.addEventListener('change', importFromExcel);

    // 添加联系方式按钮
    addMethodBtn.addEventListener('click', () => {
        // 调用A同学的函数
        addContactMethod();
    });

    // 模态框关闭按钮
    document.querySelectorAll('.close-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            contactModal.style.display = 'none';
            confirmModal.style.display = 'none';
        });
    });

    // 表单关闭按钮
    document.querySelectorAll('.close-form').forEach(btn => {
        btn.addEventListener('click', () => {
            contactModal.style.display = 'none';
        });
    });

    // 取消删除
    document.querySelector('.cancel-delete').addEventListener('click', () => {
        confirmModal.style.display = 'none';
    });

    // 确认删除
    document.querySelector('.confirm-delete').addEventListener('click', () => {
        // 调用A同学的函数
        deleteContact();
    });

    // 表单提交
    contactForm.addEventListener('submit', saveContact);

    // 点击模态框外部关闭
    window.addEventListener('click', (e) => {
        if (e.target === contactModal) contactModal.style.display = 'none';
        if (e.target === confirmModal) confirmModal.style.display = 'none';
    });
}

// 打开添加/编辑联系人模态框 - 模态框交互
function openContactModal(contact = null) {
    document.getElementById('modalTitle').textContent = contact ? '编辑联系人' : '添加联系人';
    document.getElementById('contactId').value = contact ? contact.id : '';
    document.getElementById('name').value = contact ? contact.name : '';
    document.getElementById('company').value = contact ? contact.company : '';
    document.getElementById('position').value = contact ? contact.position : '';
    document.getElementById('notes').value = contact ? contact.notes : '';
    document.getElementById('isFavorite').checked = contact ? contact.isFavorite : false;

    // 清空联系方式
    contactMethodsContainer.innerHTML = '';

    // 添加联系方式
    if (contact && contact.methods) {
        contact.methods.forEach(method => {
            // 调用A同学的函数
            addContactMethod(method.type, method.value);
        });
    } else {
        // 调用A同学的函数
        addContactMethod();
    }

    contactModal.style.display = 'flex';
}

// 保存联系人 - 表单提交处理
function saveContact(e) {
    e.preventDefault();

    const id = document.getElementById('contactId').value || generateId();
    const name = document.getElementById('name').value.trim();
    const company = document.getElementById('company').value.trim();
    const position = document.getElementById('position').value.trim();
    const notes = document.getElementById('notes').value.trim();
    const isFavorite = document.getElementById('isFavorite').checked;

    if (!name) {
        alert('请输入姓名');
        return;
    }

    // 收集联系方式
    const methods = [];
    const methodRows = contactMethodsContainer.querySelectorAll('.method-row');
    methodRows.forEach(row => {
        const type = row.querySelector('.method-type').value;
        const value = row.querySelector('.method-value').value.trim();
        if (value) {
            methods.push({ type, value });
        }
    });

    const contactData = {
        id,
        name,
        company,
        position,
        notes,
        methods,
        isFavorite,
        createdAt: new Date().toISOString()
    };

    // 更新或添加联系人
    // 需要访问A同学管理的contacts数组
    const contacts = JSON.parse(localStorage.getItem('addressBookContacts')) || [];
    const index = contacts.findIndex(c => c.id === id);
    if (index > -1) {
        contacts[index] = contactData;
    } else {
        contacts.push(contactData);
    }

    // 保存到localStorage
    localStorage.setItem('addressBookContacts', JSON.stringify(contacts));

    // 需要调用A同学的函数更新界面
    // loadContacts() 和 updateStats() 在A同学的代码中
    contactModal.style.display = 'none';
    contactForm.reset();

    // 重新加载页面以刷新联系人列表
    location.reload();
}

// 导出到Excel
function exportToExcel() {
    const contacts = JSON.parse(localStorage.getItem('addressBookContacts')) || [];

    if (contacts.length === 0) {
        alert('没有联系人可以导出');
        return;
    }

    try {
        // 创建工作表数据
        const worksheetData = [
            ['姓名', '公司', '职位', '手机', '电话', '邮箱', '微信', 'QQ', '其他联系方式', '备注', '收藏']
        ];

        contacts.forEach(contact => {
            // 初始化联系方式
            const methods = {
                '手机': '',
                '电话': '',
                '邮箱': '',
                '微信': '',
                'QQ': '',
                '其他': ''
            };

            // 填充联系方式
            contact.methods.forEach(method => {
                if (methods.hasOwnProperty(method.type)) {
                    methods[method.type] = method.value;
                }
            });

            const row = [
                contact.name || '',
                contact.company || '',
                contact.position || '',
                methods['手机'],
                methods['电话'],
                methods['邮箱'],
                methods['微信'],
                methods['QQ'],
                methods['其他'],
                contact.notes || '',
                contact.isFavorite ? '是' : '否'
            ];

            worksheetData.push(row);
        });

        // 创建工作簿
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(worksheetData);

        // 设置列宽
        const colWidths = [
            { wch: 15 }, // 姓名
            { wch: 20 }, // 公司
            { wch: 15 }, // 职位
            { wch: 15 }, // 手机
            { wch: 15 }, // 电话
            { wch: 25 }, // 邮箱
            { wch: 20 }, // 微信
            { wch: 15 }, // QQ
            { wch: 20 }, // 其他联系方式
            { wch: 30 }, // 备注
            { wch: 10 }  // 收藏
        ];
        ws['!cols'] = colWidths;

        // 添加到工作簿
        XLSX.utils.book_append_sheet(wb, ws, '通讯录');

        // 生成Excel文件并下载
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `通讯录_${new Date().toISOString().split('T')[0]}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        alert(`成功导出 ${contacts.length} 个联系人`);
    } catch (error) {
        console.error('导出失败:', error);
        alert('导出失败，请重试');
    }
}

// 从Excel导入
function importFromExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    // 检查文件类型
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        alert('请选择Excel文件 (.xlsx 或 .xls)');
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            // 跳过表头
            const rows = jsonData.slice(1);
            const importedContacts = [];

            rows.forEach((row, index) => {
                if (!row[0] || !row[0].toString().trim()) return; // 跳过空行

                const name = (row[0] || '').toString().trim();
                const company = (row[1] || '').toString().trim();
                const position = (row[2] || '').toString().trim();
                const notes = (row[9] || '').toString().trim();
                const isFavorite = (row[10] || '').toString().trim() === '是';

                // 收集联系方式
                const methods = [];

                // 检查每种联系方式列
                const methodMapping = [
                    { type: '手机', index: 3 },
                    { type: '电话', index: 4 },
                    { type: '邮箱', index: 5 },
                    { type: '微信', index: 6 },
                    { type: 'QQ', index: 7 },
                    { type: '其他', index: 8 }
                ];

                methodMapping.forEach(mapping => {
                    const value = (row[mapping.index] || '').toString().trim();
                    if (value) {
                        methods.push({ type: mapping.type, value });
                    }
                });

                // 创建联系人对象
                const contact = {
                    id: generateId(),
                    name,
                    company,
                    position,
                    notes,
                    methods,
                    isFavorite,
                    createdAt: new Date().toISOString()
                };

                importedContacts.push(contact);
            });

            if (importedContacts.length === 0) {
                alert('Excel文件中没有有效数据');
                return;
            }

            // 获取现有联系人
            const existingContacts = JSON.parse(localStorage.getItem('addressBookContacts')) || [];

            // 询问用户是覆盖还是追加
            if (existingContacts.length > 0) {
                const choice = confirm(`找到 ${importedContacts.length} 个联系人。\n点击"确定"追加到现有通讯录，点击"取消"覆盖现有通讯录。`);
                if (choice) {
                    // 追加
                    const combinedContacts = existingContacts.concat(importedContacts);
                    localStorage.setItem('addressBookContacts', JSON.stringify(combinedContacts));
                } else {
                    // 覆盖
                    localStorage.setItem('addressBookContacts', JSON.stringify(importedContacts));
                }
            } else {
                // 直接导入
                localStorage.setItem('addressBookContacts', JSON.stringify(importedContacts));
            }

            alert(`成功导入 ${importedContacts.length} 个联系人`);

            // 重新加载页面以刷新联系人列表
            location.reload();

            // 重置文件输入
            event.target.value = '';

        } catch (error) {
            console.error('导入失败:', error);
            alert('导入失败，请检查Excel文件格式是否正确');
        }
    };

    reader.readAsArrayBuffer(file);
}

// 生成唯一ID
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// 初始化代码
document.addEventListener('DOMContentLoaded', function () {
    setupEventListeners_B();
});