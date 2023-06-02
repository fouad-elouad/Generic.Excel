using System;
using System.Collections.Generic;

namespace RandomSharp
{

    public interface IRandomizer
    {

        /// <summary>
        /// Get random value from enum
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        T RandomEnum<T>() where T : struct;

        /// <summary>
        /// Get random Date between two dates
        /// </summary>
        /// <param name="min">min date</param>
        /// <param name="max">max date</param>
        /// <returns>Date</returns>
        DateTime RandomDate(DateTime min, DateTime max);

        /// <summary>
        /// Get random Date between two dates
        /// returned date can be null
        /// </summary>
        /// <param name="min">min date</param>
        /// <param name="max">max date</param>
        /// <returns>date</returns>
        DateTime? RandomNullableDate(DateTime min, DateTime max);

        /// <summary>
        /// Get random Date with time between two dates
        /// </summary>
        /// <param name="min">min datetime</param>
        /// <param name="max">max datetime</param>
        /// <returns>datetime</returns>
        DateTime RandomDateTime(DateTime min, DateTime max);

        /// <summary>
        /// Get random boolean
        /// </summary>
        /// <returns>bool</returns>
        bool RandomBool();

        /// <summary>
        /// Get random boolean
        /// </summary>
        /// <returns>bool</returns>
        bool? RandomNullableBool();

        /// <summary>
        /// Get random Int between two integers
        /// </summary>
        /// <param name="min">min int</param>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        int Random(int min, int max);

        /// <summary>
        /// Get non-negative random integer that is less than or equal to the specified maximum.
        /// </summary>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        int Random(int max);

        /// <summary>
        /// Get random Double between two doubles
        /// </summary>
        /// <param name="min">min double</param>
        /// <param name="max">max double</param>
        /// <returns>double</returns>
        double Random(double min, double max);

        /// <summary>
        /// Get random decimal between two decimals
        /// </summary>
        /// <param name="min">min decimal</param>
        /// <param name="max">max decimal</param>
        /// <returns>decimal</returns>
        decimal Random(decimal min, decimal max);

        /// <summary>
        /// Get random value from list
        /// </summary>
        /// <typeparam name="T">list type</typeparam>
        /// <param name="list">list</param>
        /// <returns>value</returns>
        T Random<T>(IList<T> list);

        /// <summary>
        /// Get random value from parameters
        /// </summary>
        /// <typeparam name="T">list type</typeparam>
        /// <param name="list">list</param>
        /// <returns>value</returns>
        T Random<T>(params T[] list) where T : struct;

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="length">length of the random string</param>
        /// <returns>string</returns>
        string RandomString(int length);

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="minLength">minimum length of the random string</param>
        /// <param name="maxLength">maximum length of the random string</param>
        /// <returns>string</returns>
        string RandomString(int minLength, int maxLength);

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        TOutput RandomNullable<TOutput>(Func<TOutput> func);

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        TOutput RandomNullable<TInput, TOutput>(TInput input, Func<TInput, TOutput> func);

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        TOutput RandomNullable<TInput1, TInput2, TOutput>(TInput1 input1, TInput2 input2, Func<TInput1, TInput2, TOutput> func);
    }
}