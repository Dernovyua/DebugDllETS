using ScriptSolution;
using ScriptSolution.Indicators;
using System;

namespace ETSdebugDll
{
    public class Class2: Script
    {
        public CreateIndicator FastMa = new CreateIndicator(EnumIndicators.MovingAvarage, 0, "�������");

        public override void Execute()
        {
            

        }

        /// <summary>
        /// ��������� ��� ����� ���� �� ������� ���� �� ���������
        /// </summary>

        public override void GetAttributesStratetgy()
        {
            DesParamStratetgy.Version = "1.0.0.5";
            DesParamStratetgy.DateRelease = "21.06.2015";
            DesParamStratetgy.DateChange = "04.04.2020";
            DesParamStratetgy.Author = "�����������";
            DesParamStratetgy.Description = "������� ������� ������� �� ���������� ����� ����� ���������, ���� � ������� �������������� �� �������� ����� (�.�. ���� � ������� �������������� �� " +
                "�������������� �����). ����� �� ������� �������������� ����� ������� �� ���������� ��������� ������ ����. ��� ����� � �������� �������: ������� ���������� ������ ����" +
                                       " ���������, � ����� �� ������� ������� �� ���������� ����� ����� ���������";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "Dll2";

        }
    }
}
