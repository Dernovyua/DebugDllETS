using ScriptSolution;
using ScriptSolution.Indicators;
using System;

namespace ETSdebugDll
{
    public class Class2: Script
    {
        public CreateIndicator FastMa = new CreateIndicator(EnumIndicators.MovingAvarage, 0, "Быстрая");

        public override void Execute()
        {
            

        }

        /// <summary>
        /// Проверить что будет если не указать один из атрибутов
        /// </summary>

        public override void GetAttributesStratetgy()
        {
            DesParamStratetgy.Version = "1.0.0.5";
            DesParamStratetgy.DateRelease = "21.06.2015";
            DesParamStratetgy.DateChange = "04.04.2020";
            DesParamStratetgy.Author = "РобоКоммерц";
            DesParamStratetgy.Description = "Условие покупки быстрая МА пересекает снизу вверх медленную, вход в позицию осуществляется на открытии свечи (т.е. вход в позицию осуществляется по " +
                "сформированным барам). Выход из позиции осуществляется когда быстрая МА пересекает медленную сверху вниз. Для входа в короткую позицию: быстрая пересекает сверху вниз" +
                                       " медленную, а выход из позиции быстрая МА пересекает снизу вверх медленную";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "Dll2";

        }
    }
}
