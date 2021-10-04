using System;
using Npoi.Mapper.Attributes;

namespace Import_Eqp
{
    ///<summary> Класс отвечает за импорт данных из Excel файла </summary>
    class Equipment
    {
        /// <summary> Наименование оборудования </summary>
        [Column(0)]
        public string FullName { get; set; }

        /// <summary> Код </summary>
        /*[Column(1)] 
        public int Code { get; set; }
        поле отсутствует в интеграционном сценарии, можно обновлять через sql запрос 
        update [dbo].[Eqp]
        Set
        [Code] = 'Код',
        [TechnicalProps] = 'Характеристики',
        [InMetrologyScope] = 1
        Where
        [MnfNum] = 'Номер производителя'
        */
        /// <summary> Технические характеристики </summary>
        /*[Ignore]
        public string TechnicalProps { get; set; }*/

        /// <summary> Наименование типа оборудования </summary> 
        [Column(2)]
        public string EqpTypeName { get; set; }

        /// <summary> Производитель </summary>
        [Column(4)]
        public string Mnf { get; set; }

        /// <summary> Заводской номер </summary>
        [Column(5)]
        public string MnfNum { get; set; }

        /// <summary> Инвентарный номер </summary>
        [Column(100)]
        public string InvNum { get; set; }

        /// <summary> Вид </summary>
        [Column(6)]
        public string EqpKindName { get; set; }

        /// <summary> Дата выпуска </summary>
        [Column(7)]
        public DateTime MnfDateStr { get; set; }

        /// <summary> Дата ввода в эксплуатацию </summary>
        [Column(8)]
        public DateTime StartUpDateStr { get; set; }

        /// <summary> Дата проведения </summary>
        [Column(9)]
        public DateTime CheckDateStr { get; set; }

        /// <summary> Состояние </summary>
        [Column(10)]
        public string EqpStateName { get; set; }

        /// <summary> Право собственности </summary>
        [Column(101)]
        public string LegalBasis { get; set; }

        /// <summary> Дата следующего проведения </summary>
        [Column(11)]
        public DateTime NextDateStr { get; set; }

        /// <summary> Примечание проверки </summary>
        [Column(15)]
        public string ECLComment { get; set; }

        /// <summary> Тип регламентных работ </summary>
        [Column(16)]
        public string EqpCheckTypeName { get; set; }

        /// <summary> Периодичность (тип интервала) </summary>
        [Column(17)]
        public string IntervalTypeName { get; set; }

        /// <summary> Периодичность (значение) </summary>
        [Column(18)]
        public int IntervalLenStr { get; set; }

        /// <summary> Назначение </summary>
        /*[Ignore]
        //public string Purpose { get; set; }*/

        /// <summary> Документ </summary>
        [Column(19)]
        public string ECLDoc { get; set; }

        /// <summary> Примечание </summary>
        [Column(20)]
        public string Note { get; set; }

        /// <summary> Лаборатория владелец </summary>
        [Column(27)]
        public string EqpSubDivisions { get; set; }

        /// <summary> Ответственный </summary>
        [Column(102)]
        public string UserId { get; set; }

        /// <summary> Место проведения проверки </summary>
        [Column(28)]
        public string ECLPlace { get; set; }

        /// <summary> Место установки </summary>
        [Column(29)]
        public string Location { get; set; }

        /// <summary> Входит в область аккредитации </summary>
        [Column(30)]
        public byte InScopeAccreditation { get; set; }

        /// <summary> Ответ службы </summary>
        [Ignore]
        public string Service_Response { get; set; }

    }
}
