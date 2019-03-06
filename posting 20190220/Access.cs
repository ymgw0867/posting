using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using MyLibrary;

namespace posting
{
    class Access
    {
        public class DataAccess
        {
            // �R���X�g���N�^
            public DataAccess()
            {
            }

            // �f�[�^���[�_�[�擾�C���^�[�t�F�C�X
            public interface IFill
            {
                // ���ۃ��\�b�h
                OleDbDataReader GetdReader(OleDbConnection tempConnection);
            }

            public class Fill��Џ�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// ��Џ��}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    mySql = "";
                    mySql += "select * from ��Џ�� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill���Ə� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// ���Ə��}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from ���Ə� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�Ј� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �Ј��}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �Ј� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �󒍃f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�󒍎�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �󒍎�ʃ}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �󒍎�� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill���� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from ���� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�U������ : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �U�������}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �U������ order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill������ : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �������f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from ������ order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�ŗ� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �ŗ��}�X�^�[���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �ŗ� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill���� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����}�X�^�[���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from ���� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�����p�^�[�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����p�^�[���}�X�^�[���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �����p�^�[�� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill���Ӑ� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// ���Ӑ�}�X�^�[���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from ���Ӑ� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill���� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from ���� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�z�z�G���A : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �z�z�G���A�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �z�z�G���A order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�z�z�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �z�z���f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �z�z�� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;

                }
            }

            public class Fill�z�z�`�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �z�z�`�ԃ}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �z�z�`�� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill�z�z�w�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �z�z�w���f�[�^�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from �z�z�w�� order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill���^ : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// ���^�}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from ���^ order by ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill�x���T�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �x���T���}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select �x���T��.*,�z�z��.���� from ";
                    mySql += "�x���T�� left join �z�z�� ";
                    mySql += "on �x���T��.�z�z��ID = �z�z��.ID ";
                    mySql += "order by �x���T��.ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill�V�� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �V��}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {

                    mySql = "";
                    mySql += "select �V��.* from �V�� ";
                    mySql += "order by �V��.���t";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill���z�z���R : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// ���z�z���R�}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {

                    mySql = "";
                    mySql += "select ���z�z���R.* from ���z�z���R ";
                    mySql += "order by ���z�z���R.ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class Fill���z�z��� : IFill
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// ���z�z���}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <returns></returns>
                public OleDbDataReader GetdReader(OleDbConnection tempConnection)
                {

                    mySql = "";
                    mySql += "select ���z�z���.* from ���z�z��� ";
                    mySql += "order by ���z�z���.ID";

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }


            // �����t���f�[�^���[�_�[�擾�C���^�[�t�F�C�X
            public interface IFillBy
            {
                // ���ۃ��\�b�h
                OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString);
            }

            // �f�[�^���[�_�[�擾�N���X
            public class free_dsReader : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL��</param>
                /// <returns>�f�[�^���[�_�[</returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += tempString;
                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;
                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy��Џ�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t����Џ��}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ��Џ�� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy���Ə� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�����Ə��}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += "select * from ���Ə� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�Ј� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���Ј��}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {

                    mySql = "";
                    mySql += "select * from �Ј� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���󒍃f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {

                    mySql = "";
                    mySql += "select * from �� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�󒍎�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���󒍎�ʃf�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from �󒍎�� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy���� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�������f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ���� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�U������ : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���U�������}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from �U������ ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy������ : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���������f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ������ ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�ŗ� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���ŗ��}�X�^�[���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from �ŗ� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy���� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�������}�X�^�[���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ���� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�����p�^�[�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�������p�^�[���}�X�^�[���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from �����p�^�[�� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy���Ӑ� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�����Ӑ�}�X�^�[���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ���Ӑ� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy���� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�������f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ���� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�z�z�G���A : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���z�z�G���A�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from �z�z�G���A ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�z�z�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���z�z���}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from �z�z�� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�z�z�`�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���z�z�`�ԃ}�X�^�[�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from �z�z�`�� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�z�z�w�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���z�z�w���f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from �z�z�w�� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy���^: IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�����^�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ���^ ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�x���T�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���x���T���f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select �x���T��.*,�z�z��.���� from ";
                    mySql += "�x���T�� left join �z�z�� ";
                    mySql += "on �x���T��.�z�z��ID = �z�z��.ID ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy�V�� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t���V��f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select �V��.* from �V�� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy���z�z���R : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�����z�z���R�f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ���z�z���R ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            public class FillBy���z�z��� : IFillBy
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                /// <summary>
                /// �����t�����z�z���f�[�^���[�_�[�擾
                /// </summary>
                /// <param name="tempConnection">�f�[�^�x�[�X�ڑ����</param>
                /// <param name="tempString">SQL���� where�ȉ��̏��������L�q���܂�</param>
                /// <returns></returns>
                public OleDbDataReader GetdsReader(OleDbConnection tempConnection, string tempString)
                {
                    mySql = "";
                    mySql += "select * from ���z�z��� ";
                    mySql += tempString;

                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;

                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

        }
    }
} 


   
