using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows;

namespace posting
{
    class Control
    {
        /// <summary>
        /// DataControl�N���X�̊�{�N���X
        /// </summary>
        public class BaseControl
        {
            private Utility.DBConnect dbConnect;

            /// <summary>
            /// �f�[�^�x�[�X�ڑ����\�b�h
            /// </summary>
            /// <returns>�f�[�^�x�[�X�ڑ������擾���܂�</returns>
            public OleDbConnection GetConnection()
            {
                //return dbConnect;
                return dbConnect.Cn;
            }

            /// <summary>
            /// BaseControl�̃R���X�g���N�^�BDBConnect�N���X�̃C���X�^���X���쐬���܂��B
            /// </summary>
            public BaseControl()
            {
                dbConnect = new Utility.DBConnect();
            }
        }

        /// <summary>
        /// �f�[�^�R���g���[���N���X BaseControl���p�����܂�
        /// </summary>
        public class DataControl : BaseControl 
        {
            private Access.DataAccess dAccess;
            public OleDbConnection cn = new OleDbConnection();

            /// <summary>
            /// DataControl�N���X�̃R���X�g���N�^�B�f�[�^�A�N�Z�X�N���X�̃C���X�^���X���쐬���܂��B
            /// </summary>
            public DataControl()
            {
                // �f�[�^�A�N�Z�X�N���X�̃C���X�^���X���쐬����
                dAccess = new Access.DataAccess();
            }

            /// <summary>
            /// �f�[�^�x�[�X�̐ڑ����������܂�
            /// </summary>
            public void Close()
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                }
            }

            /// <summary>
            /// �f�[�^���[�_�[�擾�C���^�[�t�F�C�X�������Ƃ������\�b�h
            /// </summary>
            /// <param name="IDR">�f�[�^���[�_�[���擾����C���^�[�t�F�C�X</param>
            /// <returns>�����Ŏw�肵���}�X�^�[�̃f�[�^���[�_�[</returns>
            public OleDbDataReader FillAccess(Access.DataAccess.IFill IDR)
            {
                // �f�[�^�x�[�X�ڑ������擾����
                cn = this.GetConnection();

                return IDR.GetdReader(cn);
                
            }

            /// <summary>
            /// �����t���f�[�^���[�_���擾���܂�
            /// </summary>
            /// <param name="tempString">SQL�����L�q����������</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader free_dsReader(string tempString)
            {

                try
                {
                    return FillByAccess(new Access.DataAccess.free_dsReader(), tempString);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            /// <summary>
            /// �����t���f�[�^���[�_�[�擾�C���^�[�t�F�C�X�������Ƃ������\�b�h
            /// </summary>
            /// <param name="IDSR">�f�[�^���[�_�[���擾����C���^�[�t�F�C�X</param>
            /// <param name="tempString">SQL����where�ȉ��̏������L�q����������</param>
            /// <returns>�������Ɉ�v��������Ŏw�肳�ꂽ�}�X�^�[�̃f�[�^���[�_�[</returns>
            public OleDbDataReader FillByAccess(Access.DataAccess.IFillBy IDSR,string tempString)
            {
                // �f�[�^�x�[�X�ڑ������擾����
                cn = this.GetConnection();

                return IDSR.GetdsReader(cn,tempString);

            }
        }

        /// <summary>
        /// �t���[�r�p�k���s
        /// </summary>
        public class FreeSql : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();

            public Boolean Execute(string tempSql)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    SCom.CommandText = tempSql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }
        }

        /// <summary>
        /// ��Џ��N���X
        /// </summary>
        public class ��Џ�� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// ��Џ��}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cKaisha">Entity�N���X�̉�Џ��</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.��Џ�� cKaisha)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ��Џ�� ";
                    mySql += "(��Ж�,��\�Ҏ���,��E��,�d�b�ԍ�,FAX�ԍ�,�Z��1,�Z��2,";
                    mySql += "�X�֔ԍ�,���[���A�h���X,������,�S���Җ�,���L����1,���L����2,";
                    mySql += "�˗��l�R�[�h,�˗��l��,���Z�@�փR�[�h,���Z�@�֖�,�x�X�R�[�h,�x�X��,";
                    mySql += "�������,�����ԍ�,�z�z�t���O,�o�^�N����,�ύX�N����,�X�֔ԍ�CSV�p�X, �󒍊m�菑���̓V�[�g�p�X) ";
                    mySql += "values ('" + cKaisha.��Ж� + "',";
                    mySql += "'" + cKaisha.��\�Ҏ��� + "',";
                    mySql += "'" + cKaisha.��E�� + "',";
                    mySql += "'" + cKaisha.�d�b�ԍ� + "',";
                    mySql += "'" + cKaisha.FAX�ԍ� + "',";
                    mySql += "'" + cKaisha.�Z��1 + "',";
                    mySql += "'" + cKaisha.�Z��2 + "',";
                    mySql += "'" + cKaisha.�X�֔ԍ� + "',";
                    mySql += "'" + cKaisha.���[���A�h���X + "',";
                    mySql += "'" + cKaisha.������ + "',";
                    mySql += "'" + cKaisha.�S���Җ� + "',";
                    mySql += "'" + cKaisha.���L����1 + "',";
                    mySql += "'" + cKaisha.���L����2 + "',";
                    mySql += "'" + cKaisha.�˗��l�R�[�h + "',";
                    mySql += "'" + cKaisha.�˗��l�� + "',";
                    mySql += "'" + cKaisha.���Z�@�փR�[�h + "',";
                    mySql += "'" + cKaisha.���Z�@�֖� + "',";
                    mySql += "'" + cKaisha.�x�X�R�[�h + "',";
                    mySql += "'" + cKaisha.�x�X�� + "',";
                    mySql += cKaisha.������� + ",";
                    mySql += "'" + cKaisha.�����ԍ� + "',";
                    mySql += cKaisha.�z�z�t���O + ",";
                    mySql += "'" + cKaisha.�o�^�N���� + "',";
                    mySql += "'" + cKaisha.�ύX�N���� + "',";
                    mySql += "'" + cKaisha.�X�֔ԍ�CSV�p�X + "',";
                    mySql += "'" + cKaisha.�󒍊m�菑���̓V�[�g�p�X + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ��Џ��X�V
            /// </summary>
            /// <param name="cKaisha">��Џ��Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.��Џ�� cKaisha)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ��Џ�� set ";
                    mySql += "��Ж� = '" + cKaisha.��Ж� + "',";
                    mySql += "��\�Ҏ��� = '" + cKaisha.��\�Ҏ��� + "',";
                    mySql += "��E�� = '" + cKaisha.��E�� + "',";
                    mySql += "�d�b�ԍ� = '" + cKaisha.�d�b�ԍ� + "',";
                    mySql += "FAX�ԍ� = '" + cKaisha.FAX�ԍ� + "',";
                    mySql += "�Z��1 = '" + cKaisha.�Z��1 + "',";
                    mySql += "�Z��2 = '" + cKaisha.�Z��2 + "',";
                    mySql += "�X�֔ԍ� = '" + cKaisha.�X�֔ԍ� + "',";
                    mySql += "���[���A�h���X = '" + cKaisha.���[���A�h���X + "',";
                    mySql += "������ = '" + cKaisha.������ + "',";
                    mySql += "�S���Җ� = '" + cKaisha.�S���Җ� + "',";
                    mySql += "���L����1 = '" + cKaisha.���L����1 + "',";
                    mySql += "���L����2 = '" + cKaisha.���L����2 + "',";
                    mySql += "�˗��l�R�[�h = '" + cKaisha.�˗��l�R�[�h + "',";
                    mySql += "�˗��l�� = '" + cKaisha.�˗��l�� + "',";
                    mySql += "���Z�@�փR�[�h = '" + cKaisha.���Z�@�փR�[�h + "',";
                    mySql += "���Z�@�֖� = '" + cKaisha.���Z�@�֖� + "',";
                    mySql += "�x�X�R�[�h = '" + cKaisha.�x�X�R�[�h + "',";
                    mySql += "�x�X�� = '" + cKaisha.�x�X�� + "',";
                    mySql += "������� = " + cKaisha.������� + ",";
                    mySql += "�����ԍ� = '" + cKaisha.�����ԍ� + "',";
                    mySql += "�z�z�t���O = " + cKaisha.�z�z�t���O + ",";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "',";
                    mySql += "�X�֔ԍ�CSV�p�X = '" + cKaisha.�X�֔ԍ�CSV�p�X + "',";
                    mySql += "�󒍊m�菑���̓V�[�g�p�X = '" + cKaisha.�󒍊m�菑���̓V�[�g�p�X + "' ";
                    mySql += "where ID = " + cKaisha.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ��Џ��}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ��Џ�� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ��Џ��}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>��Џ��}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // ��Џ��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill��Џ��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// ��Џ��}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // ������}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy��Џ��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// ���Ə��N���X
        /// </summary>
        public class ���Ə� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// ���Ə��}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cJigyosho">Entity�N���X�̌���</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.���Ə� cJigyosho)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ���Ə� ";
                    mySql += "(ID,����,�X�֔ԍ�,�Z��1,�Z��2,�d�b�ԍ�,FAX�ԍ�,";
                    mySql += "���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values (" + cJigyosho.ID + ",";
                    mySql += "'" + cJigyosho.���� + "',";
                    mySql += "'" + cJigyosho.�X�֔ԍ� + "',";
                    mySql += "'" + cJigyosho.�Z��1 + "',";
                    mySql += "'" + cJigyosho.�Z��2 + "',";
                    mySql += "'" + cJigyosho.�d�b�ԍ� + "',";
                    mySql += "'" + cJigyosho.FAX�ԍ� + "',";
                    mySql += "'" + cJigyosho.���l + "',";
                    mySql += "'" + cJigyosho.�o�^�N���� + "',";
                    mySql += "'" + cJigyosho.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���Ə��}�X�^�[�X�V
            /// </summary>
            /// <param name="cJigyosho">���Ə��}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.���Ə� cJigyosho)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ���Ə� set ";
                    mySql += "���� = '" + cJigyosho.���� + "',";

                    mySql += "�X�֔ԍ� = '" + cJigyosho.�X�֔ԍ� + "',";
                    mySql += "�Z��1 = '" + cJigyosho.�Z��1 + "',";
                    mySql += "�Z��2 = '" + cJigyosho.�Z��2 + "',";
                    mySql += "�d�b�ԍ� = '" + cJigyosho.�d�b�ԍ� + "',";
                    mySql += "FAX�ԍ� = '" + cJigyosho.FAX�ԍ� + "',";
                    mySql += "���l = '" + cJigyosho.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cJigyosho.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���Ə��}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ���Ə� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���Ə��}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>���Ə��}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // ���Ə��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill���Ə�());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// ���Ə��}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // ���Ə��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy���Ə�(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �󒍃N���X
        /// </summary>
        public class �� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            ///--------------------------------------------------------------------
            /// <summary>
            ///     �󒍃f�[�^�ɐV�K�Ƀf�[�^��o�^���� </summary>
            /// <param name="cJyuchu">
            ///     Entity�N���X�̎�</param>
            /// <returns>
            ///     </returns>
            ///--------------------------------------------------------------------
            public Boolean DataInsert(Entity.�� cJyuchu)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �� ";
                    mySql += "(ID,���Ə�ID,�󒍓�,�󒍋敪,���Ӑ�ID,�Ј�ID,�`���V��,�󒍎��ID,";
                    mySql += "�P��,����,���z,�����,�ō����z,�l���z,������z,�ŗ�,���^,�z�z�P��,�˗���,����,�z�z�`��,";
                    mySql += "�z�z����,�z�z�J�n��,�z�z�I����,�[�i�\���,������,";
                    mySql += "������ID,���������s��,�������@,�����\���,";
                    mySql += "�U������ID,���z�z���L��,�}�ԗL��,���L����,�G���A���l,";
                    mySql += "�����敪,���z���O,�o�^�N����,�ύX�N����,�O����ID�c��,�O���x�����c��,�O�������c��,�O���˗����c��,";
                    mySql += "�O����ID�x��,�O���x�����x��,�O�������x��,�O���˗����x��,���[�U�[ID,�Č����,";
                    mySql += "�z�z�P�\,�[�i�`��,�񍐎���,�񍐐��x,�񍐕��@,���[���A�h���X,�o�^���[�U�[ID, �O���n����,";
                    mySql += "�O���󂯓n���S����,�O���ϑ�����,�Ǝ�,";
                    mySql += "�O����ID�x��2,�O���x�����x��2,�O�������x��2,�O����ID�x��3,�O���x�����x��3,�O�������x��3,";
                    mySql += "�O���˗����x��2,�O���˗����x��3,�O���ϑ�����2,�O���ϑ�����3,�O���n����2,�O���n����3,�O���󂯓n���S����2,�O���󂯓n���S����3,";
                    mySql += "�c�Ɣ��l, �ҏW���b�N, ��������̍ς�) ";
                    mySql += "values (";
                    mySql += cJyuchu.ID + ",";
                    mySql += cJyuchu.���Ə�ID + ",";
                    mySql += "'" + cJyuchu.�󒍓� + "',";
                    mySql += "'" + cJyuchu.�󒍋敪 + "',";
                    mySql += cJyuchu.���Ӑ�ID + ",";
                    mySql += cJyuchu.�Ј�ID + ",";
                    mySql += "'" + cJyuchu.�`���V�� + "',";
                    mySql += cJyuchu.�󒍎��ID + ",";
                    mySql += cJyuchu.�P�� + ",";
                    mySql += cJyuchu.���� + ",";
                    mySql += cJyuchu.���z + ",";
                    mySql += cJyuchu.����� + ",";
                    mySql += cJyuchu.�ō����z + ",";
                    mySql += cJyuchu.�l���z + ",";
                    mySql += cJyuchu.������z + ",";
                    mySql += cJyuchu.�ŗ� + ",";
                    mySql += cJyuchu.���^ + ",";
                    mySql += cJyuchu.�z�z�P�� + ",";
                    mySql += "'" + cJyuchu.�˗��� + "',";
                    mySql += cJyuchu.���� + ",";
                    mySql += cJyuchu.�z�z�`�� + ",";
                    mySql += "'" + cJyuchu.�z�z���� + "',";

                    if (cJyuchu.�z�z�J�n�� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�z�z�J�n�� + "',";
                    }

                    if (cJyuchu.�z�z�I���� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�z�z�I���� + "',";
                    }
                    
                    if (cJyuchu.�[�i�\��� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�[�i�\��� + "',";
                    }
                    
                    mySql += cJyuchu.������ + ",";
                    mySql += cJyuchu.������ID + ",";

                    //mySql += "Null,";   //���������s��

                    // 2015/07/01
                    if (cJyuchu.���������s�� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.���������s�� + "',";
                    }
                    
                    mySql += "'" + cJyuchu.�������@ + "',";

                    if (cJyuchu.�����\��� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�����\��� + "',";
                    }
                    
                    mySql += cJyuchu.�U������ID + ",";
                    mySql += cJyuchu.���z�z���L�� + ",";
                    mySql += cJyuchu.�}�ԗL�� + ",";
                    mySql += "'" + cJyuchu.���L���� + "',";
                    mySql += "'" + cJyuchu.�G���A���l + "',";
                    mySql += cJyuchu.�����敪 + ",";
                    mySql += cJyuchu.���z���O + ",";    // 2014/11/26
                    mySql += "'" + cJyuchu.�o�^�N���� + "',";
                    mySql += "'" + cJyuchu.�ύX�N���� + "',";
                    mySql += cJyuchu.�O����ID�c��.ToString() + ",";  // 2015/06/30

                    // 2015/09/02
                    if (cJyuchu.�O����x�����c�� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O����x�����c�� + "',";
                    }

                    mySql += cJyuchu.�O���挴���c�� + ","; // 2015/06/30

                    // 2015/09/02
                    if (cJyuchu.�O����˗����c�� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O����˗����c�� + "',";
                    }

                    mySql += cJyuchu.�O����ID�x��.ToString() + ",";  // 2015/06/30

                    // 2015/09/02
                    if (cJyuchu.�O����x�����x�� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O����x�����x�� + "',";
                    }

                    mySql += cJyuchu.�O���挴���x�� + ",";     // 2015/06/30

                    // 2015/09/02
                    if (cJyuchu.�O����˗����x�� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O����˗����x�� + "',";
                    }

                    mySql += "'" + cJyuchu.���[�U�[ID + "',";   // 2015/06/30  2015/07/10
                    mySql += cJyuchu.�Č���� + ",";  // 2015/06/30
                    mySql += "'" + cJyuchu.�z�z�P�\ + "',";   // 2015/07/01
                    mySql += "'" + cJyuchu.�[�i�`�� + "',";   // 2015/07/01
                    mySql += "'" + cJyuchu.�񍐎��� + "',";     // 2015/07/01
                    mySql += "'" + cJyuchu.�񍐐��x + "',";     // 2015/07/01
                    mySql += "'" + cJyuchu.�񍐕��@ + "',";     // 2015/07/01
                    mySql += "'" + cJyuchu.���[���A�h���X + "',";  // 2015/07/01
                    mySql += "'" + cJyuchu.���[�U�[ID + "',";       // 2015/08/10

                    // 2015/09/02
                    if (cJyuchu.�O���n���� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O���n���� + "',";
                    }

                    mySql += "'" + cJyuchu.�O���󂯓n���S���� + "',";  // 2015/08/11
                    mySql += cJyuchu.�O���ϑ����� + ",";              // 2015/09/20
                    mySql += "'" + cJyuchu.�Ǝ� + "',";               // 2015/09/20
                    
                    mySql += cJyuchu.�O����ID�x��2.ToString() + ",";  // 2016/10/14

                    // 2016/10/14
                    if (cJyuchu.�O����x�����x��2 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O����x�����x��2 + "',";
                    }

                    mySql += cJyuchu.�O���挴���x��2 + ",";     // 2016/10/14

                    mySql += cJyuchu.�O����ID�x��3.ToString() + ",";  // 2016/10/14

                    // 2016/10/14
                    if (cJyuchu.�O����x�����x��3 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O����x�����x��3 + "',";
                    }

                    mySql += cJyuchu.�O���挴���x��3 + ",";     // 2016/10/14

                    // 2016/10/15
                    if (cJyuchu.�O����˗����x��2 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O����˗����x��2 + "',";
                    }

                    if (cJyuchu.�O����˗����x��3 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O����˗����x��3 + "',";
                    }

                    mySql += cJyuchu.�O���ϑ�����2 + ",";
                    mySql += cJyuchu.�O���ϑ�����3 + ",";

                    if (cJyuchu.�O���n����2 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O���n����2 + "',";
                    }

                    if (cJyuchu.�O���n����3 == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cJyuchu.�O���n����3 + "',";
                    }

                    mySql += "'" + cJyuchu.�O���󂯓n���S����2 + "',";
                    mySql += "'" + cJyuchu.�O���󂯓n���S����3 + "',";
                    mySql += "'" + cJyuchu.�c�Ɣ��l + "',";       // 2019/03/01
                    mySql += cJyuchu.�ҏW���b�N + ",";            // 2019/10/05
                    mySql += cJyuchu.��������̍ς� + ")";         // 2019/10/05

                    //System.Windows.Forms.MessageBox.Show(mySql);

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }
                catch(Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message, "�o�^�G���[", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            ///-------------------------------------------------------------
            /// <summary>
            ///     �󒍃f�[�^�X�V </summary>
            /// <param name="cJyuchu">
            ///     �󒍃f�[�^Entity�N���X</param>
            /// <returns>
            ///     �����Gtrue�A���s�Ffalse</returns>
            ///--------------------------------------------------------------
            public Boolean DataUpdate(Entity.�� cJyuchu)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �� set ";
                    mySql += "���Ə�ID = " + cJyuchu.���Ə�ID + ",";
                    mySql += "�󒍓� = '" + cJyuchu.�󒍓� + "',";
                    mySql += "�󒍋敪 = '" + cJyuchu.�󒍋敪 + "',";
                    mySql += "���Ӑ�ID = " + cJyuchu.���Ӑ�ID + ",";
                    mySql += "�Ј�ID = " + cJyuchu.�Ј�ID + ",";
                    mySql += "�`���V�� = '" + cJyuchu.�`���V�� + "',";
                    mySql += "�󒍎��ID = " + cJyuchu.�󒍎��ID + ",";
                    mySql += "�P�� = " + cJyuchu.�P�� + ",";
                    mySql += "���� = " + cJyuchu.���� + ",";
                    mySql += "���z = " + cJyuchu.���z + ",";
                    mySql += "����� = " + cJyuchu.����� + ",";
                    mySql += "�ō����z = " + cJyuchu.�ō����z + ",";
                    mySql += "�l���z = " + cJyuchu.�l���z + ",";
                    mySql += "������z = " + cJyuchu.������z + ",";
                    mySql += "�ŗ� = " + cJyuchu.�ŗ� + ",";
                    mySql += "���^ = " + cJyuchu.���^ + ",";
                    mySql += "�z�z�P�� = " + cJyuchu.�z�z�P�� + ",";
                    mySql += "�˗��� = '" + cJyuchu.�˗��� + "',";
                    mySql += "���� = " + cJyuchu.���� + ",";
                    mySql += "�z�z�`�� = " + cJyuchu.�z�z�`�� + ",";
                    mySql += "�z�z���� = '" + cJyuchu.�z�z���� + "',";

                    if (cJyuchu.�z�z�J�n�� == "")
                    {
                        mySql += "�z�z�J�n�� = Null,";
                    }
                    else
                    {
                        mySql += "�z�z�J�n�� = '" + cJyuchu.�z�z�J�n�� + "',";
                    }

                    if (cJyuchu.�z�z�I���� == "")
                    {
                        mySql += "�z�z�I���� = Null,";
                    }
                    else
                    {
                        mySql += "�z�z�I���� = '" + cJyuchu.�z�z�I���� + "',";
                    }

                    //mySql += "�z�z�P�\ = '" + cJyuchu.�z�z�P�\ + "',";    // 2015/06/23

                    if (cJyuchu.�[�i�\��� == "")
                    {
                        mySql += "�[�i�\��� = Null,";
                    }
                    else
                    {
                        mySql += "�[�i�\��� = '" + cJyuchu.�[�i�\��� + "',";
                    }

                    //mySql += "�[�i�`�� = '" + cJyuchu.�[�i�`�� + "',";    // 2015/06/23

                    mySql += "������ = " + cJyuchu.������ + ",";
                    mySql += "������ID = " + cJyuchu.������ID + ",";

                    if (cJyuchu.���������s�� == "")
                    {
                        mySql += "���������s�� = Null,";
                    }
                    else
                    {
                        mySql += "���������s�� = '" + cJyuchu.���������s�� + "',";
                    }

                    mySql += "�������@ = '" + cJyuchu.�������@ + "',";

                    if (cJyuchu.�����\��� == "")
                    {
                        mySql += "�����\��� = Null,";
                    }
                    else
                    {
                        mySql += "�����\��� = '" + cJyuchu.�����\��� + "',";
                    }

                    //mySql += "�񍐎��� = '" + cJyuchu.�񍐎��� + "',";                // 2015/06/23
                    //mySql += "�񍐐��x = '" + cJyuchu.�񍐐��x + "',";                // 2015/06/23
                    //mySql += "�񍐕��@ = '" + cJyuchu.�񍐕��@ + "',";                // 2015/06/23
                    //mySql += "���[���A�h���X = '" + cJyuchu.���[���A�h���X + "',";     // 2015/06/23

                    mySql += "�U������ID = " + cJyuchu.�U������ID + ",";
                    mySql += "���z�z���L�� = " + cJyuchu.���z�z���L�� + ",";
                    mySql += "�}�ԗL�� = " + cJyuchu.�}�ԗL�� + ",";
                    mySql += "���L���� = '" + cJyuchu.���L���� + "',";
                    mySql += "�G���A���l = '" + cJyuchu.�G���A���l + "',";
                    mySql += "�����敪 = " + cJyuchu.�����敪 + ",";
                    mySql += "���z���O = " + cJyuchu.���z���O + ",";
                    mySql += "�ύX�N���� = '" + DateTime.Now + "',";

                    mySql += "�O����ID�c�� = " + cJyuchu.�O����ID�c�� + ",";          // 2015/06/30

                    // 2015/07/17
                    if (cJyuchu.�O����x�����c�� == "")
                    {
                        mySql += "�O���x�����c�� = Null,";
                    }
                    else
                    {
                        mySql += "�O���x�����c�� = '" + cJyuchu.�O����x�����c�� + "',";
                    }

                    mySql += "�O�������c�� = " + cJyuchu.�O���挴���c�� + ",";         // 2015/06/30

                    // �ȉ��A�O����˗����c�Ǝg�p���Ȃ��̂ŃR�����g�� 2015/08/11
                    //// 2015/07/17
                    //if (cJyuchu.�O����˗����c�� == "")
                    //{
                    //    mySql += "�O���˗����c�� = Null,";
                    //}
                    //else
                    //{
                    //    mySql += "�O���˗����c�� = '" + cJyuchu.�O����˗����c�� + "',";
                    //}
                    
                    mySql += "�O����ID�x�� = " + cJyuchu.�O����ID�x�� + ",";          // 2015/06/30

                    // 2015/07/17
                    if (cJyuchu.�O����x�����x�� == "")
                    {
                        mySql += "�O���x�����x�� = Null,";
                    }
                    else
                    {
                        mySql += "�O���x�����x�� = '" + cJyuchu.�O����x�����x�� + "',";
                    }

                    mySql += "�O�������x�� = " + cJyuchu.�O���挴���x�� + ",";         // 2015/06/30

                    // 2015/07/17
                    if (cJyuchu.�O����˗����x�� == "")
                    {
                        mySql += "�O���˗����x�� = Null,";
                    }
                    else
                    {
                        mySql += "�O���˗����x�� = '" + cJyuchu.�O����˗����x�� + "',";
                    }

                    mySql += "���[�U�[ID = '" + cJyuchu.���[�U�[ID + "',";                  // 2015/06/30 2015/07/10
                    mySql += "�Č���� = " + cJyuchu.�Č���� + ",";                        // 2015/06/30

                    // 2015/08/11
                    if (cJyuchu.�O���n���� == "")
                    {
                        mySql += "�O���n���� = Null,";
                    }
                    else
                    {
                        mySql += "�O���n���� = '" + cJyuchu.�O���n���� + "',";
                    }

                    mySql += "�O���󂯓n���S���� = '" + cJyuchu.�O���󂯓n���S���� + "',";    // 2015/08/11
                    mySql += "�O���ϑ����� = " + cJyuchu.�O���ϑ����� + ",";    // 2015/09/20
                    mySql += "�Ǝ� = '" + cJyuchu.�Ǝ� + "',";                 // 2015/09/20


                    mySql += "�O����ID�x��2 = " + cJyuchu.�O����ID�x��2 + ",";          // 2016/10/14

                    // 2016/10/14
                    if (cJyuchu.�O����x�����x��2 == "")
                    {
                        mySql += "�O���x�����x��2 = Null,";
                    }
                    else
                    {
                        mySql += "�O���x�����x��2 = '" + cJyuchu.�O����x�����x��2 + "',";
                    }

                    mySql += "�O�������x��2 = " + cJyuchu.�O���挴���x��2 + ",";         // 2016/10/14


                    mySql += "�O����ID�x��3 = " + cJyuchu.�O����ID�x��3 + ",";          // 2016/10/14

                    // 2016/10/14
                    if (cJyuchu.�O����x�����x��3 == "")
                    {
                        mySql += "�O���x�����x��3 = Null,";
                    }
                    else
                    {
                        mySql += "�O���x�����x��3 = '" + cJyuchu.�O����x�����x��3 + "',";
                    }

                    mySql += "�O�������x��3 = " + cJyuchu.�O���挴���x��3 + ",";         // 2016/10/14

                    // 2016/10/15
                    if (cJyuchu.�O����˗����x��2 == "")
                    {
                        mySql += "�O���˗����x��2 = Null,";
                    }
                    else
                    {
                        mySql += "�O���˗����x��2 = '" + cJyuchu.�O����˗����x��2 + "',";
                    }

                    if (cJyuchu.�O����˗����x��3 == "")
                    {
                        mySql += "�O���˗����x��3 = Null,";
                    }
                    else
                    {
                        mySql += "�O���˗����x��3 = '" + cJyuchu.�O����˗����x��3 + "',";
                    }

                    mySql += "�O���ϑ�����2 = " + cJyuchu.�O���ϑ�����2 + ",";
                    mySql += "�O���ϑ�����3 = " + cJyuchu.�O���ϑ�����3 + ",";

                    if (cJyuchu.�O���n����2 == "")
                    {
                        mySql += "�O���n����2 = Null,";
                    }
                    else
                    {
                        mySql += "�O���n����2 = '" + cJyuchu.�O���n����2 + "',";
                    }

                    if (cJyuchu.�O���n����3 == "")
                    {
                        mySql += "�O���n����3 = Null,";
                    }
                    else
                    {
                        mySql += "�O���n����3 = '" + cJyuchu.�O���n����3 + "',";
                    }

                    mySql += "�O���󂯓n���S����2 = '" + cJyuchu.�O���󂯓n���S����2 + "',";
                    mySql += "�O���󂯓n���S����3 = '" + cJyuchu.�O���󂯓n���S����3 + "',";
                    mySql += "�c�Ɣ��l = '" + cJyuchu.�c�Ɣ��l + "',";
                    mySql += "�ҏW���b�N = " + cJyuchu.�ҏW���b�N + ",";
                    mySql += "��������̍ς� = " + cJyuchu.��������̍ς� + " ";

                    mySql += "where ID = " + cJyuchu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch(Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message, "�X�V�G���[", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            ///-----------------------------------------------------------
            /// <summary>
            ///     �󒍃f�[�^���R�[�h�폜 </summary>
            /// <param name="tempCode">
            ///     ID</param>
            /// <returns>
            ///     �����Ftrue�A���s�Ffalse</returns>
            ///-----------------------------------------------------------
            public Boolean DataDelete(long tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �󒍃f�[�^�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�󒍃f�[�^�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �󒍃f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �󒍃f�[�^�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �󒍃f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �Ј��N���X
        /// </summary>
        public class �Ј� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �Ј��}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cShain">Entity�N���X�̎Ј�</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�Ј� cShain)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �Ј� ";
                    mySql += "(ID,����,�t���K�i,�����R�[�h,��E,���ДN����,";
                    mySql += "���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values (" + cShain.ID + ",";
                    mySql += "'" + cShain.���� + "',";
                    mySql += "'" + cShain.�t���K�i + "',";
                    mySql += cShain.�����R�[�h + ",";
                    mySql += "'" + cShain.��E + "',";

                    if (cShain.���ДN���� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cShain.���ДN���� + "',";
                    }

                    mySql += "'" + cShain.���l + "',";
                    mySql += "'" + cShain.�o�^�N���� + "',";
                    mySql += "'" + cShain.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �Ј��}�X�^�[�X�V
            /// </summary>
            /// <param name="cShain">�Ј��}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�Ј� cShain)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �Ј� set ";
                    mySql += "���� = '" + cShain.���� + "',";
                    mySql += "�t���K�i = '" + cShain.�t���K�i + "',";
                    mySql += "�����R�[�h = " + cShain.�����R�[�h + ",";
                    mySql += "��E = '" + cShain.��E + "',";

                    if (cShain.���ДN���� == "")
                    {
                        mySql += "���ДN���� = Null,";
                    }
                    else
                    {
                        mySql += "���ДN���� = '" + cShain.���ДN���� + "',";
                    }

                    mySql += "���l = '" + cShain.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShain.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �Ј��}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">�Ј�ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �Ј� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �Ј��}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�Ј��}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �Ј��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�Ј�());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �Ј��}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �Ј��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�Ј�(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �󒍎�ʃN���X
        /// </summary>
        public class �󒍎�� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �󒍎�ʃ}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cJtype">Entity�N���X�̎󒍎��</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�󒍎�� cJtype)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �󒍎�� ";
                    mySql += "(����,���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cJtype.���� + "',";
                    mySql += "'" + cJtype.���l + "',";
                    mySql += "'" + cJtype.�o�^�N���� + "',";
                    mySql += "'" + cJtype.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �󒍎�ʃ}�X�^�[�X�V
            /// </summary>
            /// <param name="cJtype">�󒍎�ʃ}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�󒍎�� cJtype)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �󒍎�� set ";
                    mySql += "���� = '" + cJtype.���� + "',";
                    mySql += "���l = '" + cJtype.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cJtype.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �󒍎�ʃ}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �󒍎�� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �󒍎�ʃ}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�󒍎�ʃ}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �󒍎�ʃ}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�󒍎��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �󒍎�ʃ}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �󒍎�ʃ}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�󒍎��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �����N���X
        /// </summary>
        public class ���� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �����}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cShozoku">Entity�N���X�̏���</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.���� cShozoku)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ���� ";
                    mySql += "(ID,������1,������2,���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values (" + cShozoku.ID + ",";
                    mySql += "'" + cShozoku.������1 + "',";
                    mySql += "'" + cShozoku.������2 + "',";
                    mySql += "'" + cShozoku.���l + "',";
                    mySql += "'" + cShozoku.�o�^�N���� + "',";
                    mySql += "'" + cShozoku.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����}�X�^�[�X�V
            /// </summary>
            /// <param name="cShozoku">�����}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.���� cShozoku)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ���� set ";
                    mySql += "������1 = '" + cShozoku.������1 + "',";
                    mySql += "������2 = '" + cShozoku.������2 + "',";
                    mySql += "���l = '" + cShozoku.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShozoku.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ���� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�����}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �����}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill����());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �����}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �����}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy����(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �U�������N���X
        /// </summary>
        public class �U������ : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �U�������}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cFurikomi">Entity�N���X�̐U������</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�U������ cFurikomi)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �U������ ";
                    mySql += "(���Z�@�֖�,�x�X��,�������,�����ԍ�,�������`,�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cFurikomi.���Z�@�֖� + "',";
                    mySql += "'" + cFurikomi.�x�X�� + "',";
                    mySql += cFurikomi.������� + ",";
                    mySql += "'" + cFurikomi.�����ԍ� + "',";
                    mySql += "'" + cFurikomi.�������` + "',";
                    mySql += "'" + cFurikomi.�o�^�N���� + "',";
                    mySql += "'" + cFurikomi.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �U�������}�X�^�[�X�V
            /// </summary>
            /// <param name="cFurikomi">�U�������}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�U������ cFurikomi)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �U������ set ";
                    mySql += "���Z�@�֖� = '" + cFurikomi.���Z�@�֖� + "',";
                    mySql += "�x�X�� = '" + cFurikomi.�x�X�� + "',";
                    mySql += "������� = " + cFurikomi.������� + ",";
                    mySql += "�����ԍ� = '" + cFurikomi.�����ԍ� + "',";
                    mySql += "�������` = '" + cFurikomi.�������` + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cFurikomi.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �U�������}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �U������ ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �U�������}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�U�������}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �U�������}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�U������());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �U�������}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �U�������}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�U������(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �������N���X
        /// </summary>
        public class ������ : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �������f�[�^�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cSeikyu">Entity�N���X�̐�����</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.������ cSeikyu)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ������ ";
                    mySql += "(ID,���Ӑ�ID,�������z,�����,�l���z,������z,�ŗ�,�����\���,���s��,";
                    mySql += "�����c,�����敪,�U������ID1,�U������ID2,�o�^�N����,�ύX�N����) ";
                    mySql += "values (";
                    mySql += cSeikyu.ID + ",";
                    mySql += cSeikyu.���Ӑ�ID + ",";
                    mySql += cSeikyu.�������z + ",";
                    mySql += cSeikyu.����� + ",";
                    mySql += cSeikyu.�l���z + ",";
                    mySql += cSeikyu.������z + ",";
                    mySql += cSeikyu.�ŗ� + ",";
                    mySql += "'" + cSeikyu.�����\��� + "',";
                    mySql += "'" + cSeikyu.���s�� + "',";
                    mySql += cSeikyu.�����c + ",";
                    mySql += cSeikyu.�����敪 + ",";
                    mySql += cSeikyu.�U������ID1 + ",";
                    mySql += cSeikyu.�U������ID2 + ",";
                    mySql += "'" + cSeikyu.�o�^�N���� + "',";
                    mySql += "'" + cSeikyu.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �������f�[�^�X�V
            /// </summary>
            /// <param name="cSeikyu">������Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.������ cSeikyu)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ������ set ";
                    mySql += "���Ӑ�ID = " + cSeikyu.���Ӑ�ID + ",";
                    mySql += "�������z = " + cSeikyu.�������z + ",";
                    mySql += "����� = " + cSeikyu.����� + ",";
                    mySql += "�l���z = " + cSeikyu.�l���z + ",";
                    mySql += "������z = " + cSeikyu.������z + ",";
                    mySql += "�ŗ� = " + cSeikyu.�ŗ� + ",";
                    mySql += "�����\��� = '" + cSeikyu.�����\��� + "',";
                    mySql += "���s�� = '" + cSeikyu.���s�� + "',";
                    mySql += "�����c = " + cSeikyu.�����c + ",";
                    mySql += "�����敪 = " + cSeikyu.�����敪 + ",";
                    mySql += "�U������ID1 = " + cSeikyu.�U������ID1 + ",";
                    mySql += "�U������ID2 = " + cSeikyu.�U������ID2 + ",";
                    mySql += "���l = '" + cSeikyu.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cSeikyu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���������R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ������ ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �������f�[�^�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�������f�[�^�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �������f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill������());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �������f�[�^�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �������f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy������(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �ŗ��N���X
        /// </summary>
        public class �ŗ� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �ŗ��f�[�^�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cTax">Entity�N���X�̐ŗ�</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�ŗ� cTax)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �ŗ� ";
                    mySql += "(�ŗ�,�J�n�N����,���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values (" + cTax.�ݒ�ŗ� + ",";
                    mySql += "'" + cTax.�J�n�N���� + "',";
                    mySql += "'" + cTax.���l + "',";
                    mySql += "'" + cTax.�o�^�N���� + "',";
                    mySql += "'" + cTax.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �ŗ��}�X�^�[�X�V
            /// </summary>
            /// <param name="cTax">�ŗ�Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�ŗ� cTax)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �ŗ� set ";
                    mySql += "�ŗ� = " + cTax.�ݒ�ŗ� + ",";
                    mySql += "�J�n�N���� = '" + cTax.�J�n�N���� + "',";
                    mySql += "���l = '" + cTax.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cTax.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �ŗ����R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �ŗ� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �ŗ��}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�ŗ��}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �ŗ��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�ŗ�());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �ŗ��}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �ŗ��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�ŗ�(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �����N���X
        /// </summary>
        public class ���� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �����}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cTown">Entity�N���X�̒���</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.���� cTown)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ���� ";
                    mySql += "(ID,����,�s�撬���R�[�h,���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values (" + cTown.ID + ",";
                    mySql += "'" + cTown.���� + "',";
                    mySql += cTown.�s�撬���R�[�h + ",";
                    mySql += "'" + cTown.���l + "',";
                    mySql += "'" + cTown.�o�^�N���� + "',";
                    mySql += "'" + cTown.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����}�X�^�[�X�V
            /// </summary>
            /// <param name="cTown">����Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.���� cTown)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ���� set ";
                    mySql += "���� ='" + cTown.���� + "',";
                    mySql += "�s�撬���R�[�h = " + cTown.�s�撬���R�[�h + ",";
                    mySql += "���l ='" + cTown.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cTown.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �������R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ���� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�����}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �����}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill����());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �����}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �����}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy����(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �����p�^�[���N���X
        /// </summary>
        public class �����p�^�[�� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �����p�^�[���}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cShime">Entity�N���X�̒����p�^�[��</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�����p�^�[�� cShime)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �����p�^�[�� ";
                    mySql += "(�E�v,���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cShime.�E�v + "',";
                    mySql += "'" + cShime.���l + "',";
                    mySql += "'" + cShime.�o�^�N���� + "',";
                    mySql += "'" + cShime.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����p�^�[���}�X�^�[�X�V
            /// </summary>
            /// <param name="cShime">�����p�^�[��Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�����p�^�[�� cShime)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �����p�^�[�� set ";
                    mySql += "�E�v ='" + cShime.�E�v + "',";
                    mySql += "���l ='" + cShime.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShime.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����p�^�[�����R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �����p�^�[�� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����p�^�[���}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�����p�^�[���}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �����p�^�[���}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�����p�^�[��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �����p�^�[���}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �����p�^�[���}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�����p�^�[��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// ���Ӑ�N���X
        /// </summary>
        public class ���Ӑ� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// ���Ӑ�}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cClient">Entity�N���X�̓��Ӑ�</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.���Ӑ� cClient)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ���Ӑ� ";
                    mySql += "(����,�t���K�i,����,�h��,�S���Җ�,������,�X�֔ԍ�,�s���{��,";
                    mySql += "�Z��1,�Z��2,�d�b�ԍ�,FAX�ԍ�,���[���A�h���X,�S���Ј��R�[�h,����,";
                    mySql += "�Œʒm,������X�֔ԍ�,������s���{��,������Z��1,������Z��2,���l,";
                    mySql += "�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cClient.���� + "',";
                    mySql += "'" + cClient.�t���K�i + "',";
                    mySql += "'" + cClient.���� + "',";
                    mySql += "'" + cClient.�h�� + "',";
                    mySql += "'" + cClient.�S���Җ� + "',";
                    mySql += "'" + cClient.������ + "',";
                    mySql += "'" + cClient.�X�֔ԍ� + "',";
                    mySql += "'" + cClient.�s���{�� + "',";
                    mySql += "'" + cClient.�Z��1 + "',";
                    mySql += "'" + cClient.�Z��2 + "',";
                    mySql += "'" + cClient.�d�b�ԍ� + "',";
                    mySql += "'" + cClient.FAX�ԍ� + "',";
                    mySql += "'" + cClient.���[���A�h���X + "',";
                    mySql += cClient.�S���Ј��R�[�h + ",";
                    mySql += cClient.���� + ",";
                    mySql += "'" + cClient.�Œʒm + "',";
                    mySql += "'" + cClient.������X�֔ԍ� + "',";
                    mySql += "'" + cClient.������s���{�� + "',";
                    mySql += "'" + cClient.������Z��1 + "',";
                    mySql += "'" + cClient.������Z��2 + "',";
                    mySql += "'" + cClient.���l + "',";
                    mySql += "'" + cClient.�o�^�N���� + "',";
                    mySql += "'" + cClient.�ύX�N���� + "')";
                   
                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch(Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(e.ToString());
                    return false;
                }

            }

            /// <summary>
            /// ���Ӑ�}�X�^�[�X�V
            /// </summary>
            /// <param name="cClient">���Ӑ�Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.���Ӑ� cClient)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ���Ӑ� set ";
                    mySql += "���� = '" + cClient.���� + "',";
                    mySql += "�t���K�i = '" + cClient.�t���K�i + "',";
                    mySql += "���� = '" + cClient.���� + "',";
                    mySql += "�h�� = '" + cClient.�h�� + "',";
                    mySql += "�S���Җ� = '" + cClient.�S���Җ� + "',";
                    mySql += "������ = '" + cClient.������ + "',";
                    mySql += "�X�֔ԍ� = '" + cClient.�X�֔ԍ� + "',";
                    mySql += "�s���{�� = '" + cClient.�s���{�� + "',";
                    mySql += "�Z��1 = '" + cClient.�Z��1 + "',";
                    mySql += "�Z��2 = '" + cClient.�Z��2 + "',";
                    mySql += "�d�b�ԍ� = '" + cClient.�d�b�ԍ� + "',";
                    mySql += "FAX�ԍ� = '" + cClient.FAX�ԍ� + "',";
                    mySql += "���[���A�h���X = '" + cClient.���[���A�h���X + "',";
                    mySql += "�S���Ј��R�[�h = " + cClient.�S���Ј��R�[�h + ",";
                    mySql += "���� = " + cClient.���� + ",";
                    mySql += "�Œʒm = '" + cClient.�Œʒm + "',";
                    mySql += "������X�֔ԍ� = '" + cClient.������X�֔ԍ� + "',";
                    mySql += "������s���{�� = '" + cClient.������s���{�� + "',";
                    mySql += "������Z��1 = '" + cClient.������Z��1 + "',";
                    mySql += "������Z��2 = '" + cClient.������Z��2 + "',";
                    mySql += "���l ='" + cClient.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cClient.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���Ӑ惌�R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ���Ӑ� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���Ӑ�}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>���Ӑ�}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // ���Ӑ�}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill���Ӑ�());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// ���Ӑ�}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // ���Ӑ�}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy���Ӑ�(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �����N���X
        /// </summary>
        public class ���� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �����f�[�^�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cNyukin">Entity�N���X�̓���</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.���� cNyukin)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ���� ";
                    mySql += "(������ID,�����N����,���z,���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values (" + cNyukin.������ID + ",";
                    mySql += "'" + cNyukin.�����N���� + "',";
                    mySql += cNyukin.���z + ",";
                    mySql += "'" + cNyukin.���l + "',";
                    mySql += "'" + cNyukin.�o�^�N���� + "',";
                    mySql += "'" + cNyukin.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����}�X�^�[�X�V
            /// </summary>
            /// <param name="cNyukin">����Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.���� cNyukin)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ���� set ";
                    mySql += "������ID = " + cNyukin.������ID + ",";
                    mySql += "�����N���� = '" + cNyukin.�����N���� + "',";
                    mySql += "���z = " + cNyukin.���z + ",";
                    mySql += "���l ='" + cNyukin.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cNyukin.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �������R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ���� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �����f�[�^�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�����f�[�^�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �����f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill����());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �����f�[�^�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �����f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy����(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �z�z�G���A�N���X
        /// </summary>
        public class �z�z�G���A : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �z�z�G���A�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cArea">Entity�N���X�̔z�z�G���A</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�z�z�G���A cArea)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �z�z�G���A ";
                    mySql += "(����ID,�\�薇��,��ID,�z�z�w��ID,�z�z�P��,�z�z��,���z�z����,���c��,";
                    mySql += "�񍐖���,�񍐎c��,���z�敪,�}�ԋL��,�����敪,�X�e�[�^�X,�o�^�N����,�ύX�N����) ";
                    mySql += "values (" + cArea.����ID + ",";
                    mySql += cArea.�\�薇�� + ",";
                    mySql += cArea.��ID+ ",";
                    mySql += cArea.�z�z�w��ID + ",";
                    mySql += cArea.�z�z�P�� + ",";

                    if (cArea.�z�z�� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cArea.�z�z�� + "',";
                    }

                    mySql += cArea.���z�z���� + ",";
                    mySql += cArea.���c�� + ",";
                    mySql += cArea.�񍐖��� + ",";
                    mySql += cArea.�񍐎c�� + ",";
                    mySql += cArea.���z�敪 + ",";
                    mySql += "'" + cArea.�}�ԋL�� + "',";
                    mySql += cArea.�����敪 + ",";
                    mySql += cArea.�X�e�[�^�X + ",";
                    mySql += "'" + cArea.�o�^�N���� + "',";
                    mySql += "'" + cArea.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�G���A�X�V
            /// </summary>
            /// <param name="cArea">�z�z�G���AEntity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�z�z�G���A cArea)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �z�z�G���A set ";
                    mySql += "����ID = " + cArea.����ID + ",";
                    mySql += "�\�薇�� = " + cArea.�\�薇�� + ",";
                    mySql += "��ID = " + cArea.��ID+ ",";
                    mySql += "�z�z�w��ID = " + cArea.�z�z�w��ID + ",";
                    mySql += "�z�z�P�� = " + cArea.�z�z�P�� + ",";


                    if (cArea.�z�z�� == "")
                    {
                        mySql += "�z�z�� = Null,";
                    }
                    else
                    {
                        mySql += "�z�z�� = '" + cArea.�z�z�� + "',";
                    }

                    mySql += "���z�z���� = " + cArea.���z�z���� + ",";
                    mySql += "���c�� = " + cArea.���c�� + ",";
                    mySql += "�񍐖��� = " + cArea.�񍐖��� + ",";
                    mySql += "�񍐎c�� = " + cArea.�񍐎c�� + ",";
                    mySql += "���z�敪 = " + cArea.���z�敪 + ",";
                    mySql += "�}�ԋL�� = '" + cArea.�}�ԋL�� + "',";
                    mySql += "�����敪 = " + cArea.�����敪 + ",";
                    mySql += "�X�e�[�^�X = " + cArea.�X�e�[�^�X + ",";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cArea.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�G���A���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �z�z�G���A ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�G���A�f�[�^�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�����f�[�^�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �z�z�G���A�f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�z�z�G���A());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �z�z�G���A�f�[�^�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �z�z�G���A�f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�z�z�G���A(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// ���z�z���N���X
        /// </summary>
        public class ���z�z��� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// ���z�z���ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cMihaifu">Entity�N���X�̖��z�z���</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.���z�z��� cMihaifu)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ���z�z��� ";
                    mySql += "(�z�z�G���AID,�Ԓn��,�}���V������,���R,���̑����e,�o�^�N����,�ύX�N����) ";
                    mySql += "values (" + cMihaifu.�z�z�G���AID + ",";
                    mySql += "'" + cMihaifu.�Ԓn��  + "',";
                    mySql += "'" + cMihaifu.�}���V������ + "',";
                    mySql += cMihaifu.���R + ",";
                    mySql += "'" + cMihaifu.���̑����e + "',";
                    mySql += "'" + cMihaifu.�o�^�N���� + "',";
                    mySql += "'" + cMihaifu.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���z�z���X�V
            /// </summary>
            /// <param name="cMihaifu">���z�z���Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.���z�z��� cMihaifu)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ���z�z��� set ";
                    mySql += "�z�z�G���AID = " + cMihaifu.�z�z�G���AID + ",";
                    mySql += "�Ԓn�� = '" + cMihaifu.�Ԓn�� + "',";
                    mySql += "�}���V������ = '" + cMihaifu.�}���V������ + "',";
                    mySql += "���R = " + cMihaifu.���R + ",";
                    mySql += "���̑����e = '" + cMihaifu.���̑����e + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cMihaifu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���z�z��񃌃R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ���z�z��� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���z�z���f�[�^�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>���z�z���S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // ���z�z���f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill���z�z���());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// ���z�z���f�[�^�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // ���z�z���f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy���z�z���(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }


        /// <summary>
        /// �z�z���N���X
        /// </summary>
        public class �z�z�� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// -------------------------------------------------------------------------
            /// <summary>
            ///     �z�z���}�X�^�[�ɐV�K�Ƀf�[�^��o�^���� </summary>
            /// <param name="cStaff">
            ///     Entity�N���X�̔z�z��</param>
            /// <returns>
            ///     �o�^����:true, �o�^���s:false</returns>
            ///     
            ///     2015/07/16 �}�C�i���o�[,���[�U�[ID�ǉ�
            /// -------------------------------------------------------------------------
            public Boolean DataInsert(Entity.�z�z�� cStaff)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �z�z�� ";
                    mySql += "(ID,����,�t���K�i,�X�֔ԍ�,�Z��,�g�ѓd�b�ԍ�,����d�b�ԍ�,PC�A�h���X,";
                    mySql += "�g�уA�h���X,�o�^��,�Ζ��敪,�X���z�z�敪,�X���z�z���l,�x���敪,";
                    mySql += "���Ə��R�[�h,���Z�@�փR�[�h,���Z�@�֖�,���Z�@�֖��J�i,�x�X�R�[�h,�x�X��,�x�X���J�i,";
                    mySql += "�������,�����ԍ�,�������`�J�i,���l,";
                    mySql += "�o�^�N����,�ύX�N����,�}�C�i���o�[,���[�U�[ID) ";
                    mySql += "values (" + cStaff.ID + ",";
                    mySql += "'" + cStaff.���� + "',";
                    mySql += "'" + cStaff.�t���K�i + "',";
                    mySql += "'" + cStaff.�X�֔ԍ� + "',";
                    mySql += "'" + cStaff.�Z�� + "',";
                    mySql += "'" + cStaff.�g�ѓd�b�ԍ� + "',";
                    mySql += "'" + cStaff.����d�b�ԍ� + "',";
                    mySql += "'" + cStaff.PC�A�h���X + "',";
                    mySql += "'" + cStaff.�g�уA�h���X + "',";

                    if (cStaff.�o�^�� == "")
                    {
                        mySql += "Null,";
                    }
                    else
                    {
                        mySql += "'" + cStaff.�o�^�� + "',";
                    }

                    mySql += cStaff.�Ζ��敪 + ",";
                    mySql += cStaff.�X���z�z�敪 + ",";
                    mySql += "'" + cStaff.�X���z�z���l + "',";
                    mySql += "'" + cStaff.�x���敪 + "',";
                    mySql += cStaff.���Ə��R�[�h + ",";
                    mySql += "'" + cStaff.���Z�@�փR�[�h + "',";
                    mySql += "'" + cStaff.���Z�@�֖� + "',";
                    mySql += "'" + cStaff.���Z�@�֖��J�i + "',";
                    mySql += "'" + cStaff.�x�X�R�[�h + "',";
                    mySql += "'" + cStaff.�x�X�� + "',";
                    mySql += "'" + cStaff.�x�X���J�i + "',";
                    mySql += cStaff.������� + ",";
                    mySql += "'" + cStaff.�����ԍ� + "',";
                    mySql += "'" + cStaff.�������`�J�i + "',";
                    mySql += "'" + cStaff.���l + "',";
                    mySql += "'" + cStaff.�o�^�N���� + "',";
                    mySql += "'" + cStaff.�ύX�N���� + "',";
                    mySql += "'" + cStaff.�}�C�i���o�[ + "',";    // 2015/07/16
                    mySql += cStaff.���[�U�[ID + ")";             // 2015/07/16

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            /// -------------------------------------------------------------------------
            /// <summary>
            ///     �z�z���X�V </summary>
            /// <param name="cStaff">
            ///     �z�z�G���AEntity�N���X</param>
            /// <returns>
            ///     �����Gtrue�A���s�Ffalse</returns>
            ///     
            ///     2015/07/16 �}�C�i���o�[,���[�U�[ID�ǉ�
            /// -------------------------------------------------------------------------
            public Boolean DataUpdate(Entity.�z�z�� cStaff)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �z�z�� set ";
                    mySql += "���� = '" + cStaff.���� + "',";
                    mySql += "�t���K�i = '" + cStaff.�t���K�i + "',";
                    mySql += "�X�֔ԍ� = '" + cStaff.�X�֔ԍ� + "',";
                    mySql += "�Z�� = '" + cStaff.�Z�� + "',";
                    mySql += "�g�ѓd�b�ԍ� = '" + cStaff.�g�ѓd�b�ԍ� + "',";
                    mySql += "����d�b�ԍ� = '" + cStaff.����d�b�ԍ� + "',";
                    mySql += "PC�A�h���X = '" + cStaff.PC�A�h���X + "',";
                    mySql += "�g�уA�h���X = '" + cStaff.�g�уA�h���X + "',";

                    if (cStaff.�o�^�� == "")
                    {
                        mySql += "�o�^�� = Null,";
                    }
                    else
                    {
                        mySql += "�o�^�� = '" + cStaff.�o�^�� + "',";
                    }

                    mySql += "�Ζ��敪 = " + cStaff.�Ζ��敪 + ",";
                    mySql += "�X���z�z�敪 = " + cStaff.�X���z�z�敪 + ",";
                    mySql += "�X���z�z���l = '" + cStaff.�X���z�z���l + "',";
                    mySql += "�x���敪 = '" +cStaff.�x���敪 + "',";
                    mySql += "���Ə��R�[�h = " + cStaff.���Ə��R�[�h + ",";
                    mySql += "���Z�@�փR�[�h = '" + cStaff.���Z�@�փR�[�h + "',";
                    mySql += "���Z�@�֖� = '" + cStaff.���Z�@�֖� + "',";
                    mySql += "���Z�@�֖��J�i = '" + cStaff.���Z�@�֖��J�i + "',";
                    mySql += "�x�X�R�[�h = '" + cStaff.�x�X�R�[�h + "',";
                    mySql += "�x�X�� = '" + cStaff.�x�X�� + "',";
                    mySql += "�x�X���J�i = '" + cStaff.�x�X���J�i + "',";
                    mySql += "������� = " + cStaff.������� + ",";
                    mySql += "�����ԍ� = '" + cStaff.�����ԍ� + "',";
                    mySql += "�������`�J�i = '" + cStaff.�������`�J�i + "',";
                    mySql += "���l = '" + cStaff.���l + "',";
                    mySql += "�ύX�N���� = '" + cStaff.�ύX�N���� + "',";
                    mySql += "�}�C�i���o�[ = '" + cStaff.�}�C�i���o�[ + "',";   // 2015/07/16
                    mySql += "���[�U�[ID = " + cStaff.���[�U�[ID + " ";         // 2015/07/16
                    mySql += "where ID = " + cStaff.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            ///-------------------------------------------------------------------
            /// <summary>
            ///     �z�z�����R�[�h�폜 </summary>
            /// <param name="tempCode">
            ///     ID</param>
            /// <returns>
            ///     �����Ftrue�A���s�Ffalse</returns>
            ///-------------------------------------------------------------------
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �z�z�� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            /// <summary>
            /// �z�z���}�X�^�[�f�[�^�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�z�z���}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �z�z���}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�z�z��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �z�z���}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �z�z���}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�z�z��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �z�z�`�ԃN���X
        /// </summary>
        public class �z�z�`�� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �z�z�`�ԃ}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cHaifu">Entity�N���X�̔z�z�`��</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�z�z�`�� cHaifu)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �z�z�`�� ";
                    mySql += "(����,��l�����薇��,���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cHaifu.���� + "',";
                    mySql += cHaifu.��l�����薇�� + ",";
                    mySql += "'" + cHaifu.���l + "',";
                    mySql += "'" + cHaifu.�o�^�N���� + "',";
                    mySql += "'" + cHaifu.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�`�ԃ}�X�^�[�X�V
            /// </summary>
            /// <param name="cHaifu">����Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�z�z�`�� cHaifu)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �z�z�`�� set ";
                    mySql += "���� = '" + cHaifu.���� + "',";
                    mySql += "��l�����薇�� = " + cHaifu.��l�����薇�� + ",";
                    mySql += "���l ='" + cHaifu.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cHaifu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�`�ԃ��R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �z�z�`�� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�`�ԃ}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�z�z�`�ԃ}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �z�z�`�ԃ}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�z�z�`��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �z�z�`�ԃ}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �z�z�`�ԃ}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�z�z�`��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// �z�z�w���N���X
        /// </summary>
        public class �z�z�w�� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            ///------------------------------------------------------------------------------
            /// <summary>
            ///     �z�z�w���f�[�^�ɐV�K�Ƀf�[�^��o�^���� </summary>
            /// <param name="cShiji">
            ///     Entity�N���X�̔z�z�w��</param>
            /// <returns>
            ///     </returns>
            ///------------------------------------------------------------------------------
            public Boolean DataInsert(Entity.�z�z�w�� cShiji)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �z�z�w�� ";
                    mySql += "(ID,�z�z��,���͓�,�z�z��ID,��ʔ�,��ʋ�ԊJ�n,��ʋ�ԏI��,";
                    mySql += "�z�z�J�n����,�z�z�I������,�I�����|�[�g,���z�z�敪,���z�z���R,";
                    mySql += "���ӎ���,";
                    mySql += "�o�^�N����,�ύX�N����) ";
                    mySql += "values (";
                    mySql += cShiji.ID + ",";
                    
                    //if (cShiji.�z�z�� == "")
                    //{
                    //    mySql += "Null,";
                    //}
                    //else
                    //{
                    //    mySql += "'" + cShiji.�z�z�� + "',";
                    //}

                    mySql += "'" + cShiji.�z�z�� + "',";
                    mySql += "'" + cShiji.���͓� + "',";
                    mySql += cShiji.�z�z��ID + ",";
                    mySql += cShiji.��ʔ� + ",";
                    mySql += "'" + cShiji.��ʋ�ԊJ�n + "',";
                    mySql += "'" + cShiji.��ʋ�ԏI�� + "',";
                    mySql += "'" + cShiji.�z�z�J�n���� + "',";
                    mySql += "'" + cShiji.�z�z�I������ + "',";
                    mySql += "'" + cShiji.�I�����|�[�g + "',";
                    mySql += "'" + cShiji.���z�z�敪 + "',";
                    mySql += "'" + cShiji.���z�z���R + "',";
                    mySql += "'" + cShiji.���ӎ��� + "',";
                    mySql += "'" + cShiji.�o�^�N���� + "',";
                    mySql += "'" + cShiji.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�w���f�[�^�X�V
            /// </summary>
            /// <param name="cShiji">�z�z�w���f�[�^Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�z�z�w�� cShiji)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �z�z�w�� set ";

                    //if (cShiji.�z�z�� == "")
                    //{
                    //    mySql += "�z�z�� = Null,";
                    //}
                    //else
                    //{
                    //    mySql += "�z�z�� = '" + cShiji.�z�z�� + "',";
                    //}

                    mySql += "�z�z�� = '" + cShiji.�z�z�� + "',";
                    mySql += "���͓� = '" + cShiji.���͓� + "',";
                    mySql += "�z�z��ID = " + cShiji.�z�z��ID + ",";
                    mySql += "��ʔ� = " + cShiji.��ʔ� + ",";
                    mySql += "��ʋ�ԊJ�n = '" + cShiji.��ʋ�ԊJ�n + "',";
                    mySql += "��ʋ�ԏI�� = '" + cShiji.��ʋ�ԏI�� + "',";
                    mySql += "�z�z�J�n���� = '" + cShiji.�z�z�J�n���� + "',";
                    mySql += "�z�z�I������ = '" + cShiji.�z�z�I������ + "',";
                    mySql += "�I�����|�[�g = '" + cShiji.�I�����|�[�g + "',";
                    mySql += "���z�z�敪 = '" + cShiji.���z�z�敪 + "',";
                    mySql += "���z�z���R = '" + cShiji.���z�z���R + "',";
                    mySql += "���ӎ��� = '" + cShiji.���ӎ��� + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShiji.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�w�����R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �z�z�w�� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �z�z�w���f�[�^�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�z�z�w���f�[�^�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // �z�z�w���f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�z�z�w��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �z�z�w���f�[�^�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �z�z�w���f�[�^�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�z�z�w��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// ���^�N���X
        /// </summary>
        public class ���^ : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// ���^�}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cSize">Entity�N���X�̔��^</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.���^ cSize)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ���^ ";
                    mySql += "(����,���P��1,���P��2,���P��3,���l,�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cSize.���� + "',";
                    mySql += cSize.���P��1 + ",";
                    mySql += cSize.���P��2 + ",";
                    mySql += cSize.���P��3 + ",";
                    mySql += "'" + cSize.���l + "',";
                    mySql += "'" + cSize.�o�^�N���� + "',";
                    mySql += "'" + cSize.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���^�}�X�^�[�X�V
            /// </summary>
            /// <param name="cSize">���^�}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.���^ cSize)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ���^ set ";
                    mySql += "���� = '" + cSize.���� + "',";
                    mySql += "���P��1 = " + cSize.���P��1 + ",";
                    mySql += "���P��2 = " + cSize.���P��2 + ",";
                    mySql += "���P��3 = " + cSize.���P��3 + ",";
                    mySql += "���l = '" + cSize.���l + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cSize.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���^�}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ���^ ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���^�}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>���^�}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // ���^�}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill���^());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// ���^�}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // ���^�}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy���^(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// �x���T���N���X
        /// </summary>
        public class �x���T�� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �x���T���}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cShikyu">Entity�N���X�̔��^</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�x���T�� cShikyu)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �x���T�� ";
                    mySql += "(���t,�z�z��ID,�E�v,�P��,����,���z,�x���T���敪,�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cShikyu.���t + "',";
                    mySql += cShikyu.�z�z��ID + ",";
                    mySql += "'" + cShikyu.�E�v + "',";
                    mySql += cShikyu.�P�� + ",";
                    mySql += cShikyu.���� + ",";
                    mySql += cShikyu.���z + ",";
                    mySql += cShikyu.�x���T���敪 + ",";
                    mySql += "'" + cShikyu.�o�^�N���� + "',";
                    mySql += "'" + cShikyu.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �x���T���}�X�^�[�X�V
            /// </summary>
            /// <param name="cShikyu">�x���T���}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�x���T�� cShikyu)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �x���T�� set ";
                    mySql += "���t = '" + cShikyu.���t + "',";
                    mySql += "�z�z��ID = " + cShikyu.�z�z��ID + ",";
                    mySql += "�E�v = '" + cShikyu.�E�v + "',";
                    mySql += "�P�� = " + cShikyu.�P�� + ",";
                    mySql += "���� = " + cShikyu.���� + ",";
                    mySql += "���z = " + cShikyu.���z + ",";
                    mySql += "�x���T���敪 = " + cShikyu.�x���T���敪 + ",";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cShikyu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �x���T���}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempCode">ID</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempCode)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �x���T�� ";
                    mySql += "where ID = " + tempCode;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �x���T���}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�x���T���}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    // ���^�}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�x���T��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �x���T���}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �x���T���}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�x���T��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        public class �V�� : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// �x���T���}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cShikyu">Entity�N���X�̔��^</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.�V�� cTenkou)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into �V�� ";
                    mySql += "(���t,�V��,�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cTenkou.���t + "',";
                    mySql += "'" + cTenkou.�V�� + "',";
                    mySql += "'" + cTenkou.�o�^�N���� + "',";
                    mySql += "'" + cTenkou.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �V��}�X�^�[�X�V
            /// </summary>
            /// <param name="cShikyu">�V��}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.�V�� cTenkou)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update �V�� set ";
                    mySql += "�V�� = '" + cTenkou.�V�� + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ���t = '" + cTenkou.���t + "'";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �V��}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempDate">���t</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(DateTime tempDate)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from �V�� ";
                    mySql += "where ���t = " + tempDate;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// �V��}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>�V��}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    //�V��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill�V��());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// �V��}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �V��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy�V��(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        public class ���z�z���R : DataControl
        {
            private static OleDbCommand SCom = new OleDbCommand();
            private static String mySql;

            /// <summary>
            /// ���z�z���R�}�X�^�[�ɐV�K�Ƀf�[�^��o�^����
            /// </summary>
            /// <param name="cRiyu">Entity�N���X�̖��z�z���R</param>
            /// <returns></returns>
            public Boolean DataInsert(Entity.���z�z���R cRiyu)
            {

                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    //�o�^����
                    mySql = "";
                    mySql += "insert into ���z�z���R ";
                    mySql += "(ID,�E�v,�o�^�N����,�ύX�N����) ";
                    mySql += "values ('" + cRiyu.ID + "',";
                    mySql += "'" + cRiyu.�E�v + "',";
                    mySql += "'" + cRiyu.�o�^�N���� + "',";
                    mySql += "'" + cRiyu.�ύX�N���� + "')";

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���z�z���R�}�X�^�[�X�V
            /// </summary>
            /// <param name="cRiyu">���z�z���R�}�X�^�[Entity�N���X</param>
            /// <returns>�����Gtrue�A���s�Ffalse</returns>
            public Boolean DataUpdate(Entity.���z�z���R cRiyu)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "update ���z�z���R set ";
                    mySql += "�E�v = '" + cRiyu.�E�v + "',";
                    mySql += "�ύX�N���� = '" + DateTime.Today + "' ";
                    mySql += "where ID = " + cRiyu.ID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���z�z���R�}�X�^�[���R�[�h�폜
            /// </summary>
            /// <param name="tempDate">���t</param>
            /// <returns>�����Ftrue�A���s�Ffalse</returns>
            public Boolean DataDelete(int tempID)
            {
                try
                {
                    //�f�[�^�x�[�X�ڑ����̎擾
                    cn = this.GetConnection();

                    mySql = "";
                    mySql += "delete from ���z�z���R ";
                    mySql += "where ID = " + tempID;

                    SCom.CommandText = mySql;
                    SCom.Connection = cn;

                    // SQL�̎��s
                    SCom.ExecuteNonQuery();

                    return true;
                }
                catch
                {
                    return false;
                }

            }

            /// <summary>
            /// ���z�z���R�}�X�^�[�S���̃f�[�^���[�_���擾���܂�
            /// </summary>
            /// <returns>���z�z���R�}�X�^�[�S���f�[�^���[�_�[</returns>
            public OleDbDataReader Fill()
            {
                try
                {
                    //���z�z���R�}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillAccess(new Access.DataAccess.Fill���z�z���R());
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            /// <summary>
            /// ���z�z���R�}�X�^�[�����t�f�[�^���[�_�[�擾
            /// </summary>
            /// <param name="tempString">�������FSQL����where��ȉ����L�q���܂�</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public OleDbDataReader FillBy(string tempString)
            {
                try
                {
                    // �V��}�X�^�[�̃f�[�^���[�_�[�擾�N���X�̃C���X�^���X�������Ŏw��
                    return FillByAccess(new Access.DataAccess.FillBy���z�z���R(), tempString);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

    }
}
