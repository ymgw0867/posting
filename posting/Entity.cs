using System;
using System.Collections.Generic;
using System.Text;

namespace posting
{
    class Entity
    {
        public class ��Џ��
        {
            public int ID { get; set; }
            public string ��Ж� { get; set; }
            public string ��\�Ҏ��� { get; set; }
            public string ��E�� { get; set; }
            public string �d�b�ԍ� { get; set; }
            public string FAX�ԍ� { get; set; }
            public string �Z��1 { get; set; }
            public string �Z��2 { get; set; }
            public string �X�֔ԍ� { get; set; }
            public string ���[���A�h���X { get; set; }
            public string ������ { get; set; }
            public string �S���Җ� { get; set; }
            public string ���L����1 { get; set; }
            public string ���L����2 { get; set; }
            public string �˗��l�R�[�h { get; set; }
            public string �˗��l�� { get; set; }
            public string ���Z�@�փR�[�h { get; set; }
            public string ���Z�@�֖� { get; set; }
            public string �x�X�R�[�h { get; set; }
            public string �x�X�� { get; set; }
            public int ������� { get; set; }
            public string �����ԍ� { get; set; }
            public int �z�z�t���O { get; set; }
            public DateTime �o�^�N���� { get; set; }
            public DateTime �ύX�N���� { get; set; }
            public string �X�֔ԍ�CSV�p�X { get; set; } 
            public string �󒍊m�菑���̓V�[�g�p�X { get; set; }    // 2019/03/06
        }

        public class ���Ə�
        {
            private int F_ID;
            private string F_����;
            private string F_�X�֔ԍ�;
            private string F_�Z��1;
            private string F_�Z��2;
            private string F_�d�b�ԍ�;
            private string F_FAX�ԍ�;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string ����
            {
                set
                {
                    F_���� = value;
                }
                get
                {
                    return F_����;
                }
            }

            public string �X�֔ԍ�
            {
                set
                {
                    F_�X�֔ԍ� = value;
                }
                get
                {
                    return F_�X�֔ԍ�;
                }
            }

            public string �Z��1
            {
                set
                {
                    F_�Z��1 = value;
                }
                get
                {
                    return F_�Z��1;
                }
            }

            public string �Z��2
            {
                set
                {
                    F_�Z��2 = value;
                }
                get
                {
                    return F_�Z��2;
                }
            }

            public string �d�b�ԍ�
            {
                set
                {
                    F_�d�b�ԍ� = value;
                }
                get
                {
                    return F_�d�b�ԍ�;
                }
            }

            public string FAX�ԍ�
            {
                set
                {
                    F_FAX�ԍ� = value;
                }
                get
                {
                    return F_FAX�ԍ�;
                }
            }

            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }
            
            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class �Ј�
        {
            private int F_ID;
            private string F_����;
            private string F_�t���K�i;
            private int F_�����R�[�h;
            private string F_��E;
            private string F_���ДN����;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;
 
            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string ����
            {
                set
                {
                    F_���� = value;
                }
                get
                {
                    return F_����;
                }
            }

            public string �t���K�i
            {
                set
                {
                    F_�t���K�i = value;
                }
                get
                {
                    return F_�t���K�i;
                }
            }

            public int �����R�[�h
            {
                set
                {
                    F_�����R�[�h = value;
                }
                get
                {
                    return F_�����R�[�h;
                }
            }

            public string ��E
            {
                set
                {
                    F_��E = value;
                }
                get
                {
                    return F_��E;
                }
            }

            public string ���ДN����
            {
                set
                {
                    F_���ДN���� = value;
                }
                get
                {
                    return F_���ДN����;
                }
            }

            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }

            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class ��
        {
            public long ID { get; set; }            
            public int ���Ə�ID { get; set; }            
            public DateTime �󒍓� { get; set; }            
            public string �󒍋敪 { get; set; }            
            public int ���Ӑ�ID { get; set; }            
            public int �Ј�ID { get; set; }            
            public string �`���V�� { get; set; }            
            public int �󒍎��ID { get; set; }            
            public double �P�� { get; set; }            
            public int ���� { get; set; }            
            public int ���z { get; set; }            
            public int ����� { get; set; }            
            public int �ō����z { get; set; }            
            public int �l���z { get; set; }            
            public int ������z { get; set; }            
            public int �ŗ� { get; set; }            
            public int ���^ { get; set; }            
            public double �z�z�P�� { get; set; }            
            public string �˗��� { get; set; }            
            public double ���� { get; set; }            
            public int �z�z�`�� { get; set; }            
            public string �z�z���� { get; set; }            
            public string �z�z�J�n�� { get; set; }            
            public string �z�z�I���� { get; set; }            
            public string �z�z�P�\ { get; set; }            
            public string �[�i�\��� { get; set; }            
            public string �[�i�`�� { get; set; }            
            public int ������ { get; set; }            
            public int ������ID { get; set; }            
            public string ���������s�� { get; set; }            
            public string �������@ { get; set; }            
            public string �����\��� { get; set; }            
            public string �񍐎��� { get; set; }            
            public string �񍐐��x { get; set; }            
            public string �񍐕��@ { get; set; }            
            public string ���[���A�h���X { get; set; }            
            public int �U������ID { get; set; }            
            public int ���z�z���L�� { get; set; }            
            public int �}�ԗL�� { get; set; }            
            public string ���L���� { get; set; }          
            public string �G���A���l { get; set; }
            public int �����敪 { get; set; }
            public DateTime �o�^�N���� { get; set; }
            public DateTime �ύX�N���� { get; set; }
            public long IDTemplateS { get; set; }
            public long IDTemplateE { get; set; }
            public int ���z���O { get; set; }

            // 2015/06/30 �ǉ��t�B�[���h
            public int �O����ID�c�� { get; set; }
            public string �O����x�����c�� { get; set; }
            public double �O���挴���c�� { get; set; }
            public string �O����˗����c�� { get; set; }    // 2015/07/17

            public int �O����ID�x�� { get; set; }
            public string �O����x�����x�� { get; set; }
            public double �O���挴���x�� { get; set; }
            public int �O����ID�x��2 { get; set; }           // 2016/10/14
            public string �O����x�����x��2 { get; set; }   // 2016/10/14
            public double �O���挴���x��2 { get; set; }    // 2016/10/14
            public int �O����ID�x��3 { get; set; }       // 2016/10/14
            public string �O����x�����x��3 { get; set; }   // 2016/10/14
            public double �O���挴���x��3 { get; set; }    // 2016/10/14
            public string �O����˗����x�� { get; set; }    // 2015/07/17
            public string �O����˗����x��2 { get; set; }    // 2016/10/15
            public string �O����˗����x��3 { get; set; }    // 2016/10/15

            public int ���[�U�[ID { get; set; }
            public int �Č���� { get; set; }
            public string �O���n���� { get; set; }           // 2015/08/11
            public string �O���n����2 { get; set; }           // 2016/10/15
            public string �O���n����3 { get; set; }           // 2016/10/15
            public string �O���󂯓n���S���� { get; set; }   // 2015/08/11
            public string �O���󂯓n���S����2 { get; set; }   // 2016/10/15
            public string �O���󂯓n���S����3 { get; set; }   // 2016/10/15

            public int �O���ϑ����� { get; set; }     // 2015/09/20
            public int �O���ϑ�����2 { get; set; }     // 2016/10/15
            public int �O���ϑ�����3 { get; set; }     // 2016/10/15
            public string �Ǝ� { get; set; }         // 2015/09/20      
            public string �c�Ɣ��l { get; set; }   // 2019/03/01
        }

        public class �󒍎��
        {
            private int F_ID;
            private string F_����;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string ����
            {
                set
                {
                    F_���� = value;
                }
                get
                {
                    return F_����;
                }
            }

            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }

            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class ����
        {
            private int F_ID;
            private string F_������1;
            private string F_������2;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string ������1
            {
                set
                {
                    F_������1 = value;
                }
                get
                {
                    return F_������1;
                }
            }

            public string ������2
            {
                set
                {
                    F_������2 = value;
                }
                get
                {
                    return F_������2;
                }
            }

            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }

            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class �U������
        {
            private int F_ID;
            private string F_���Z�@�֖�;
            private string F_�x�X��;
            private int F_�������;
            private string F_�����ԍ�;
            private string F_�������`;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }

            public string ���Z�@�֖�
            {
                get
                {
                    return F_���Z�@�֖�;
                }
                set
                {
                    F_���Z�@�֖� = value;
                }
            }

            public string �x�X��
            {
                get
                {
                    return F_�x�X��;
                }
                set
                {
                    F_�x�X�� = value;
                }
            }

            public int �������
            {
                get
                {
                    return F_�������;
                }
                set
                {
                    F_������� = value;
                }
            }

            public string �����ԍ�
            {
                get
                {
                    return F_�����ԍ�;
                }
                set
                {
                    F_�����ԍ� = value;
                }
            }

            public string �������`
            {
                get
                {
                    return F_�������`;
                }
                set
                {
                    F_�������` = value;
                }
            }

            public DateTime �o�^�N����
            {
                get
                {
                    return F_�o�^�N����;
                }
                set
                {
                    F_�o�^�N���� = value;
                }
            }

            public DateTime �ύX�N����
            {
                get
                {
                    return F_�ύX�N����;
                }
                set
                {
                    F_�ύX�N���� = value;
                }
            }
       }
       
        public class ������
        {
            private int F_ID;
            private int F_���Ӑ�ID;
            private int F_�������z;
            private int F_�����;
            private int F_�l���z;
            private int F_������z;
            private int F_�ŗ�;
            private DateTime F_�����\���;
            private DateTime F_���s��;
            private int F_�����c;
            private int F_�����敪;
            private int F_�U������ID1;
            private int F_�U������ID2;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                get
                {
                    return F_ID;
                }
                set
                {
                    F_ID = value;
                }
            }
            public int ���Ӑ�ID
            {
                get
                {
                    return F_���Ӑ�ID;
                }
                set
                {
                    F_���Ӑ�ID = value;
                }
            }
            public int �������z
            {
                get
                {
                    return F_�������z;
                }
                set
                {
                    F_�������z = value;
                }
            }
            public int �����
            {
                get
                {
                    return F_�����;
                }
                set
                {
                    F_����� = value;
                }
            }
            public int �l���z
            {
                get
                {
                    return F_�l���z;
                }
                set
                {
                    F_�l���z = value;
                }
            }

            public int ������z
            {
                get
                {
                    return F_������z;
                }
                set
                {
                    F_������z = value;
                }
            }
            public int �ŗ�
            {
                get
                {
                    return F_�ŗ�;
                }
                set
                {
                    F_�ŗ� = value;
                }
            }
            public DateTime �����\���
            {
                get
                {
                    return F_�����\���;
                }
                set
                {
                    F_�����\��� = value;
                }
            }
            public DateTime ���s��
            {
                get
                {
                    return F_���s��;
                }
                set
                {
                    F_���s�� = value;
                }
            }
            public int �����c
            {
                get
                {
                    return F_�����c;
                }
                set
                {
                    F_�����c = value;
                }
            }
            public int �����敪
            {
                get
                {
                    return F_�����敪;
                }
                set
                {
                    F_�����敪 = value;
                }
            }
            public int �U������ID1
            {
                get
                {
                    return F_�U������ID1;
                }
                set
                {
                    F_�U������ID1 = value;
                }
            }
            public int �U������ID2
            {
                get
                {
                    return F_�U������ID2;
                }
                set
                {
                    F_�U������ID2 = value;
                }
            }
            public string ���l
            {
                get
                {
                    return F_���l;
                }
                set
                {
                    F_���l = value;
                }
            }
            public DateTime �o�^�N����
            {
                get
                {
                    return F_�o�^�N����;
                }
                set
                {
                    F_�o�^�N���� = value;
                }
            }
            public DateTime �ύX�N����
            {
                get
                {
                    return F_�ύX�N����;
                }
                set
                {
                    F_�ύX�N���� = value;
                }
            }
        }

        public class �ŗ�
        {
            private int F_ID;
            private int F_�ŗ�;
            private DateTime F_�J�n�N����;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public int �ݒ�ŗ�
            {
                set
                {
                    F_�ŗ� = value;
                }
                get
                {
                    return F_�ŗ�;
                }
            }

            public DateTime �J�n�N����
            {
                set
                {
                    F_�J�n�N���� = value;
                }
                get
                {
                    return F_�J�n�N����;
                }
            }

            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }

            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class ����
        {
            private int F_ID;
            private string F_����;
            private int F_�s�撬���R�[�h;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string ����
            {
                set
                {
                    F_���� = value;
                }
                get
                {
                    return F_����;
                }
            }

            public int �s�撬���R�[�h
            {
                get { return F_�s�撬���R�[�h; }
                set { F_�s�撬���R�[�h = value; }
            }

            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }

            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class �����p�^�[��
        {
            private int F_ID;
            private string F_�E�v;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string �E�v
            {
                set
                {
                    F_�E�v = value;
                }
                get
                {
                    return F_�E�v;
                }
            }

            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }

            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class ���Ӑ�
        {
            public int ID { set; get; }
            public string ���� { set; get; }
            public string �t���K�i { set; get; }
            public string ���� { set; get; }   
            public string �h�� { set; get; }
            public string �S���Җ� { set; get; }
            public string ������ { set; get; }
            public string �X�֔ԍ� { set; get; }
            public string �s���{�� { set; get; }
            public string �Z��1 { set; get; }
            public string �Z��2 { set; get; }
            public string �d�b�ԍ� { set; get; }
            public string FAX�ԍ� { set; get; }
            public string ���[���A�h���X { set; get; }
            public int �S���Ј��R�[�h { set; get; }
            public int ���� { set; get; }
            public string �Œʒm { set; get; }
            public string ������X�֔ԍ� { set; get; }
            public string ������s���{�� { set; get; }
            public string ������Z��1 { set; get; }
            public string ������Z��2 { set; get; }
            public string ���l { set; get; }
            public DateTime �o�^�N���� { set; get; }
            public DateTime �ύX�N���� { set; get; }
            public string �����於�� { set; get; }
            public string ������S���Җ� { set; get; }
        }

        public class ����
        {
            private int F_ID;
            private int F_������ID;
            private DateTime F_�����N����;
            private int F_���z;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public int ������ID
            {
                set
                {
                    F_������ID = value;
                }
                get
                {
                    return F_������ID;
                }
            }

            public DateTime �����N����
            {
                set
                {
                    F_�����N���� = value;
                }
                get
                {
                    return F_�����N����;
                }
            }

            public int ���z
            {
                set
                {
                    F_���z = value;
                }
                get
                {
                    return F_���z;
                }
            }

            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }

            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class �z�z�G���A
        {
            private int F_ID;
            private int F_����ID;
            private int F_�\�薇��;
            private long F_��ID;
            private int F_�z�z�w��ID;
            private double F_�z�z�P��;
            private string F_�z�z��;
            private int F_���z�z����;
            private int F_���c��;
            private int F_�񍐖���;
            private int F_�񍐎c��;
            private int F_���z�敪;
            private string F_�}�ԋL��;
            private int F_�����敪;
            private int F_�X�e�[�^�X;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }
            public int ����ID
            {
                set
                {
                    F_����ID = value;
                }
                get
                {
                    return F_����ID;
                }
            }
            public int �\�薇��
            {
                set
                {
                    F_�\�薇�� = value;
                }
                get
                {
                    return F_�\�薇��;
                }
            }
            public long ��ID
            {
                set
                {
                    F_��ID = value;
                }
                get
                {
                    return F_��ID;
                }
            }
            public int �z�z�w��ID
            {
                set
                {
                    F_�z�z�w��ID = value;
                }
                get
                {
                    return F_�z�z�w��ID;
                }
            }
            public double �z�z�P��
            {
                set
                {
                    F_�z�z�P�� = value;
                }
                get
                {
                    return F_�z�z�P��;
                }
            }
            public string �z�z��
            {
                set
                {
                    F_�z�z�� = value;
                }
                get
                {
                    return F_�z�z��;
                }
            }
            public int ���z�z����
            {
                set
                {
                    F_���z�z���� = value;
                }
                get
                {
                    return F_���z�z����;
                }
            }
            public int ���c��
            {
                set
                {
                    F_���c�� = value;
                }
                get
                {
                    return F_���c��;
                }
            }
            public int �񍐖���
            {
                set
                {
                    F_�񍐖��� = value;
                }
                get
                {
                    return F_�񍐖���;
                }
            }
            public int �񍐎c��
            {
                set
                {
                    F_�񍐎c�� = value;
                }
                get
                {
                    return F_�񍐎c��;
                }
            }
            public int ���z�敪
            {
                set
                {
                    F_���z�敪 = value;
                }
                get
                {
                    return F_���z�敪;
                }
            }
            public string �}�ԋL��
            {
                set
                {
                    F_�}�ԋL�� = value;
                }
                get
                {
                    return F_�}�ԋL��;
                }
            }
            public int �����敪
            {
                set
                {
                    F_�����敪 = value;
                }
                get
                {
                    return F_�����敪;
                }
            }
            public int �X�e�[�^�X
            {
                set
                {
                    F_�X�e�[�^�X = value;
                }
                get
                {
                    return F_�X�e�[�^�X;
                }
            }
            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }
            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class ���z�z���
        {
            private int F_ID;

            public int ID
            {
                get { return F_ID; }
                set { F_ID = value; }
            }

            private int F_�z�z�G���AID;

            public int �z�z�G���AID
            {
                get { return F_�z�z�G���AID; }
                set { F_�z�z�G���AID = value; }
            }

            private string F_�Ԓn��;

            public string �Ԓn��
            {
                get { return F_�Ԓn��; }
                set { F_�Ԓn�� = value; }
            }

            private string F_�}���V������;

            public string �}���V������
            {
                get { return F_�}���V������; }
                set { F_�}���V������ = value; }
            }

            private int F_���R;

            public int ���R
            {
                get { return F_���R; }
                set { F_���R = value; }
            }

            private string F_���̑����e;

            public string ���̑����e
            {
                get { return F_���̑����e; }
                set { F_���̑����e = value; }
            }

            private DateTime F_�o�^�N����;

            public DateTime �o�^�N����
            {
                get { return F_�o�^�N����; }
                set { F_�o�^�N���� = value; }
            }

            private DateTime F_�ύX�N����;

            public DateTime �ύX�N����
            {
                get { return F_�ύX�N����; }
                set { F_�ύX�N���� = value; }
            }

        }

        public class �z�z��
        {
            public int ID { set; get; }
            public string ���� { set; get; }
            public string �t���K�i { set; get; }
            public string �X�֔ԍ� { set; get; }
            public string �Z�� { set; get; }
            public string �g�ѓd�b�ԍ� { set; get; }
            public string ����d�b�ԍ� { set; get; }
            public string PC�A�h���X { set; get; }
            public string �g�уA�h���X { set; get; }
            public string �o�^�� { set; get; }
            public int �Ζ��敪 { set; get; }
            public int �X���z�z�敪 { set; get; }
            public string �X���z�z���l { set; get; }
            public string �x���敪 { set; get; }
            public int ���Ə��R�[�h { set; get; }
            public string ���Z�@�փR�[�h { set; get; }
            public string ���Z�@�֖� { set; get; }
            public string ���Z�@�֖��J�i { set; get; }
            public string �x�X�R�[�h { set; get; }
            public string �x�X�� { set; get; }
            public string �x�X���J�i { set; get; }
            public int ������� { set; get; }
            public string �����ԍ� { set; get; }
            public string �������`�J�i { set; get; }
            public string ���l { set; get; }
            public DateTime �o�^�N���� { set; get; }
            public DateTime �ύX�N���� { set; get; }            
            public int ���[�U�[ID { get; set; }
            public string �}�C�i���o�[ { get; set; }            
        }

        public class �z�z�`��
        {
            public int ID { set; get; }
            public string ���� { set; get; }
            public int ��l�����薇�� { set; get; }
            public string ���l { set; get; }
            public DateTime �o�^�N���� { set; get; }
            public DateTime �ύX�N���� { set; get; }            
        }

        public class �z�z�w��
        {
            private int F_ID;
            private DateTime F_�z�z��;
            private DateTime F_���͓�;
            private int F_�z�z��ID;
            private int F_��ʔ�;
            private string F_��ʋ�ԊJ�n;
            private string F_��ʋ�ԏI��;
            private string F_�z�z�J�n����;
            private string F_�z�z�I������;
            private string F_�I�����|�[�g;
            private string F_���z�z�敪;
            private string F_���z�z���R;
            private string F_���ӎ���;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;
            private int F_���[�U�[ID;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public DateTime �z�z��
            {
                set
                {
                    F_�z�z�� = value;
                }
                get
                {
                    return F_�z�z��;
                }
            }

            public DateTime ���͓�
            {
                set
                {
                    F_���͓� = value;
                }
                get
                {
                    return F_���͓�;
                }
            }

            public int �z�z��ID
            {
                set
                {
                    F_�z�z��ID = value;
                }
                get
                {
                    return F_�z�z��ID;
                }
            }

            public int ��ʔ�
            {
                set
                {
                    F_��ʔ� = value;
                }
                get
                {
                    return F_��ʔ�;
                }
            }

            public string ��ʋ�ԊJ�n
            {
                set
                {
                    F_��ʋ�ԊJ�n = value;
                }
                get
                {
                    return F_��ʋ�ԊJ�n;
                }
            }

            public string ��ʋ�ԏI��
            {
                set
                {
                    F_��ʋ�ԏI�� = value;
                }
                get
                {
                    return F_��ʋ�ԏI��;
                }
            }

            public string �z�z�J�n����
            {
                set
                {
                    F_�z�z�J�n���� = value;
                }
                get
                {
                    return F_�z�z�J�n����;
                }
            }

            public string �z�z�I������
            {
                set
                {
                    F_�z�z�I������ = value;
                }
                get
                {
                    return F_�z�z�I������;
                }
            }

            public string �I�����|�[�g
            {
                set
                {
                    F_�I�����|�[�g = value;
                }
                get
                {
                    return F_�I�����|�[�g;
                }
            }

            public string ���z�z�敪
            {
                set
                {
                    F_���z�z�敪 = value;
                }
                get
                {
                    return F_���z�z�敪;
                }
            }

            public string ���z�z���R
            {
                set
                {
                    F_���z�z���R = value;
                }
                get
                {
                    return F_���z�z���R;
                }
            }

            public string ���ӎ���
            {
                set
                {
                    F_���ӎ��� = value;
                }
                get
                {
                    return F_���ӎ���;
                }
            }


            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
            public int ���[�U�[ID
            {
                set
                {
                    F_���[�U�[ID = value;
                }
                get
                {
                    return F_���[�U�[ID;
                }
            }
        }

        public class ���^
        {
            private int F_ID;
            private string F_����;
            private double F_���P��1;
            private double F_���P��2;
            private double F_���P��3;
            private string F_���l;
            private DateTime F_�o�^�N����;
            private DateTime F_�ύX�N����;

            public int ID
            {
                set
                {
                    F_ID = value;
                }
                get
                {
                    return F_ID;
                }
            }

            public string ����
            {
                set
                {
                    F_���� = value;
                }
                get
                {
                    return F_����;
                }
            }

            public double ���P��1
            {
                set
                {
                    F_���P��1 = value;
                }
                get
                {
                    return F_���P��1;
                }
            }

            public double ���P��2
            {
                set
                {
                    F_���P��2 = value;
                }
                get
                {
                    return F_���P��2;
                }
            }

            public double ���P��3
            {
                set
                {
                    F_���P��3 = value;
                }
                get
                {
                    return F_���P��3;
                }
            }
            
            public string ���l
            {
                set
                {
                    F_���l = value;
                }
                get
                {
                    return F_���l;
                }
            }

            public DateTime �o�^�N����
            {
                set
                {
                    F_�o�^�N���� = value;
                }
                get
                {
                    return F_�o�^�N����;
                }
            }

            public DateTime �ύX�N����
            {
                set
                {
                    F_�ύX�N���� = value;
                }
                get
                {
                    return F_�ύX�N����;
                }
            }
        }

        public class �x���T��
        {
            //ID
            private int F_ID;

            public int ID
            {
                get { return F_ID; }
                set { F_ID = value; }
            }
	
            //���t
            private DateTime F_���t;

            public DateTime ���t
            {
                get { return F_���t; }
                set { F_���t = value; }
            }

            private int F_�z�z��ID;

            public int �z�z��ID
            {
                get { return F_�z�z��ID; }
                set { F_�z�z��ID = value; }
            }

            private string F_�z�z����;

            public string �z�z����
            {
                get { return F_�z�z����; }
                set { F_�z�z���� = value; }
            }
	
            private string F_�E�v;

            public string �E�v
            {
                get { return F_�E�v; }
                set { F_�E�v = value; }
            }

            private double F_�P��;

            public double �P��
            {
                get { return F_�P��; }
                set { F_�P�� = value; }
            }

            private int F_����;

            public int ����
            {
                get { return F_����; }
                set { F_���� = value; }
            }

            private double F_���z;

            public double ���z
            {
                get { return F_���z; }
                set { F_���z = value; }
            }

            private int F_�x���T���敪;

            public int �x���T���敪
            {
                get { return F_�x���T���敪; }
                set { F_�x���T���敪 = value; }
            }

            private DateTime F_�o�^�N����;

            public DateTime �o�^�N����
            {
                get { return F_�o�^�N����; }
                set { F_�o�^�N���� = value; }
            }

            private DateTime F_�ύX�N����;

            public DateTime �ύX�N����
            {
                get { return F_�ύX�N����; }
                set { F_�ύX�N���� = value; }
            }
	
        }

        public class �V��
        {
            private DateTime F_���t;

            public DateTime ���t
            {
                get { return F_���t; }
                set { F_���t = value; }
            }

            private string F_�V��;

            public string �V��
            {
                get { return F_�V��; }
                set { F_�V�� = value; }
            }

            private DateTime F_�o�^�N����;

            public DateTime �o�^�N����
            {
                get { return F_�o�^�N����; }
                set { F_�o�^�N���� = value; }
            }

            private DateTime F_�ύX�N����;

            public DateTime �ύX�N����
            {
                get { return F_�ύX�N����; }
                set { F_�ύX�N���� = value; }
            }
        }

        public class ���z�z���R
        {
            private int F_ID;

            public int ID
            {
                get { return F_ID; }
                set { F_ID = value; }
            }

            private string F_�E�v;

            public string �E�v
            {
                get { return F_�E�v; }
                set { F_�E�v = value; }
            }

            private DateTime F_�o�^�N����;

            public DateTime �o�^�N����
            {
                get { return F_�o�^�N����; }
                set { F_�o�^�N���� = value; }
            }

            private DateTime F_�ύX�N����;

            public DateTime �ύX�N����
            {
                get { return F_�ύX�N����; }
                set { F_�ύX�N���� = value; }
            }
        }

        public class �S��w�b�_
        {
            private string F_�f�[�^�敪;

            public string �f�[�^�敪
            {
                get { return F_�f�[�^�敪; }
                set { F_�f�[�^�敪 = value; }
            }

            private string F_��ʃR�[�h;

            public string ��ʃR�[�h
            {
                get { return F_��ʃR�[�h; }
                set { F_��ʃR�[�h = value; }
            }

            private string F_�R�[�h�敪;

            public string �R�[�h�敪
            {
                get { return F_�R�[�h�敪; }
                set { F_�R�[�h�敪 = value; }
            }

            private string F_�U���˗��l�R�[�h;

            public string �U���˗��l�R�[�h
            {
                get { return F_�U���˗��l�R�[�h; }
                set { F_�U���˗��l�R�[�h = value; }
            }

            private string F_�U���˗��l��;

            public string �U���˗��l��
            {
                get { return F_�U���˗��l��; }
                set { F_�U���˗��l�� = value; }
            }

            private string F_��g��;

            public string ��g��
            {
                get { return F_��g��; }
                set { F_��g�� = value; }
            }

            private string F_�d����s�ԍ�;

            public string �d����s�ԍ�
            {
                get { return F_�d����s�ԍ�; }
                set { F_�d����s�ԍ� = value; }
            }

            private string F_�d����s��;

            public string �d����s��
            {
                get { return F_�d����s��; }
                set { F_�d����s�� = value; }
            }

            private string F_�d���x�X�ԍ�;

            public string �d���x�X�ԍ�
            {
                get { return F_�d���x�X�ԍ�; }
                set { F_�d���x�X�ԍ� = value; }
            }

            private string F_�d���x�X��;

            public string �d���x�X��
            {
                get { return F_�d���x�X��; }
                set { F_�d���x�X�� = value; }
            }

            private string F_�a�����;

            public string �a�����
            {
                get { return F_�a�����; }
                set { F_�a����� = value; }
            }

            private string F_�����ԍ�;

            public string �����ԍ�
            {
                get { return F_�����ԍ�; }
                set { F_�����ԍ� = value; }
            }

            private string F_�_�~�[;

            public string �_�~�[
            {
                get { return F_�_�~�[; }
                set { F_�_�~�[ = value; }
            }
	
        }

        public class �S��f�[�^���R�[�h
        {
            private string F_�f�[�^�敪;

            public string �f�[�^�敪
            {
                get { return F_�f�[�^�敪; }
                set { F_�f�[�^�敪 = value; }
            }

            private string F_��d����s�ԍ�;

            public string ��d����s�ԍ�
            {
                get { return F_��d����s�ԍ�; }
                set { F_��d����s�ԍ� = value; }
            }

            private string F_��d����s��;

            public string ��d����s��
            {
                get { return F_��d����s��; }
                set { F_��d����s�� = value; }
            }

            private string F_��d���x�X�ԍ�;

            public string ��d���x�X�ԍ�
            {
                get { return F_��d���x�X�ԍ�; }
                set { F_��d���x�X�ԍ� = value; }
            }

            private string F_��d���x�X��;

            public string ��d���x�X��
            {
                get { return F_��d���x�X��; }
                set { F_��d���x�X�� = value; }
            }

            private string F_��`�������ԍ�;

            public string ��`�������ԍ�
            {
                get { return F_��`�������ԍ�; }
                set { F_��`�������ԍ� = value; }
            }

            private string F_�a�����;

            public string �a�����
            {
                get { return F_�a�����; }
                set { F_�a����� = value; }
            }

            private string F_�����ԍ�;

            public string �����ԍ�
            {
                get { return F_�����ԍ�; }
                set { F_�����ԍ� = value; }
            }

            private string F_���l��;

            public string ���l��
            {
                get { return F_���l��; }
                set { F_���l�� = value; }
            }

            private string F_�U�����z;

            public string �U�����z
            {
                get { return F_�U�����z; }
                set { F_�U�����z = value; }
            }

            private string F_�V�K�R�[�h;

            public string �V�K�R�[�h
            {
                get { return F_�V�K�R�[�h; }
                set { F_�V�K�R�[�h = value; }
            }

            private string F_�ڋq�R�[�h1;

            public string �ڋq�R�[�h1
            {
                get { return F_�ڋq�R�[�h1; }
                set { F_�ڋq�R�[�h1 = value; }
            }

            private string F_�ڋq�R�[�h2;

            public string �ڋq�R�[�h2
            {
                get { return F_�ڋq�R�[�h2; }
                set { F_�ڋq�R�[�h2 = value; }
            }

            private string F_EDI���;

            public string EDI���
            {
                get { return F_EDI���; }
                set { F_EDI��� = value; }
            }

            private string F_�U���w��敪;

            public string �U���w��敪
            {
                get { return F_�U���w��敪; }
                set { F_�U���w��敪 = value; }
            }

            private string F_���ʕ\��;

            public string ���ʕ\��
            {
                get { return F_���ʕ\��; }
                set { F_���ʕ\�� = value; }
            }

            private string F_�_�~�[;

            public string �_�~�[
            {
                get { return F_�_�~�[; }
                set { F_�_�~�[ = value; }
            }
	
        }

        public class �S��g���[���[���R�[�h
        {
            private string F_�f�[�^�敪;

            public string �f�[�^�敪
            {
                get { return F_�f�[�^�敪; }
                set { F_�f�[�^�敪 = value; }
            }

            private string F_���v����;

            public string ���v����
            {
                get { return F_���v����; }
                set { F_���v���� = value; }
            }

            private string F_���v���z;

            public string ���v���z
            {
                get { return F_���v���z; }
                set { F_���v���z = value; }
            }

            private string F_�_�~�[;

            public string �_�~�[
            {
                get { return F_�_�~�[; }
                set { F_�_�~�[ = value; }
            }
        }

        public class �S��G���h���R�[�h
        {
            private string F_�f�[�^�敪;

            public string �f�[�^�敪
            {
                get { return F_�f�[�^�敪; }
                set { F_�f�[�^�敪 = value; }
            }

            private string F_�_�~�[;

            public string �_�~�[
            {
                get { return F_�_�~�[; }
                set { F_�_�~�[ = value; }
            }
        }

        //�ėp�f�[�^�w�b�_����
        public class OutPutHeader
        {
            public const string dn01 = @"""OBCD001""";  // �`�[���

            public const string hd00 = @"""CSJS003""";  // ����w����@ 
            public const string hd01 = @"""CSJS004""";  // �`�[����R�[�h 
            public const string hd02 = @"""CSJS005""";  // ���t 
            public const string hd03 = @"""CSJS007""";  // �`�[��

            public const string kr01 = @"""CSJS200""";  // �ؕ�����R�[�h
            public const string kr02 = @"""CSJS201""";  // �ؕ�����ȖڃR�[�h
            public const string kr03 = @"""CSJS202""";  // �ؕ��⏕�ȖڃR�[�h
            public const string kr04 = @"""CSJS205""";  // �ؕ����Ƌ敪�R�[�h
            public const string kr05 = @"""CSJS207""";  // �ؕ��[������
            public const string kr55 = @"""CSJS208""";  // �ؕ������R�[�h
            public const string kr06 = @"""CSJS213""";  // �ؕ��{�̋��z

            public const string ks01 = @"""CSJS300""";  // �ݕ�����R�[�h
            public const string ks02 = @"""CSJS301""";  // �ݕ�����ȖڃR�[�h
            public const string ks52 = @"""CSJS302""";  // �ݕ��⏕�ȖڃR�[�h�E�E�E�Œ�l�i0�j 2019/09/27
            public const string ks53 = @"""CSJS304""";  // �ݕ��ŗ��敪�R�[�h
            public const string ks03 = @"""CSJS305""";  // �ݕ����Ƌ敪�R�[�h
            public const string ks55 = @"""CSJS306""";  // �ݕ�����Ōv�Z
            public const string ks04 = @"""CSJS307""";  // �ݕ��[������
            public const string ks05 = @"""CSJS308""";  // �ݕ������R�[�h
            public const string ks06 = @"""CSJS313""";  // �ݕ��{�̋��z
            public const string ks54 = @"""CSJS320""";  // �ݕ��ŗ�
            public const string ks56 = @"""CSJS322""";  // �ŗ���� �E�E�E�Œ�l�i0�F�W���j 2019/09/27

            public const string tk01 = @"""CSJS100""";  // �E�v
            
        }


    }
}
