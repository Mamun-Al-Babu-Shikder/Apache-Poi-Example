package com.mcubes;

import java.util.HashMap;
import java.util.Map;

public class ConfidenceIntervalCalculator {

    public static void main(String[] args) {
        ConfidenceIntervalCalculator cal = new ConfidenceIntervalCalculator();
        Map<String, Double> m = cal.getLowerAndUpperBound(95, 849000, 385, 20);
        System.out.println(m);
        System.out.println(cal.hypergeometric_confidence_interval(849000, 385, 20, 95));
    }


    public Map<String, Double> getLowerAndUpperBound(double confidenceLevel, long populationSize, long sampleSize,
                                                     long relevantDocs) {
        return hypergeometric_confidence_interval(populationSize, sampleSize, relevantDocs, confidenceLevel);
    }

    private double log_n_choose_m(double n, double m) {
        double result = 0.0;
        double max = n - m > m ? n - m : m;
        double j = n - max;
        for (double i = max + 1; i <= n; ++i) {
            if (j > 1)
            {
                result += Math.log(i / j);
                --j;
            }
            else
                result += Math.log(i);
        }
        while (j > 1) {
            result -= Math.log(j);
            --j;
        }
        return result;
    }

    private double hypergeometric_prob(double N, double M, double n, double m) {
        if (M == 0)
            return m == 0 ? 1.0 : 0.0;
        if (M == N)
            return m == n ? 1.0 : 0.0;

        if (M < m)
            return 0.0;
        if (N - M < n - m)
            return 0.0;

        double i_n = N - n + 1;
        double i_m = M - m + 1;
        double i_n_minus_m = N - M - (n - m) + 1;

        double log_p_part = 0.0;

        double n_minus_m = N - M;
        while (i_n_minus_m <= n_minus_m) {
            log_p_part += Math.log(i_n_minus_m / i_n);
            ++i_n_minus_m;
            ++i_n;
        }
        while (i_m <= M) {
            log_p_part += Math.log(i_m / i_n);
            ++i_m;
            ++i_n;
        }
        return Math.exp(log_n_choose_m(n, m) + log_p_part);
    }

    private double hypergeometric_cumulative_prob(double N, double M, double n, double m) {
        double result = 0.0;

        double i_max = m;
        if (i_max > M)
            i_max = M;

        double i_min = 0;
        if (i_min < n - (N - M))
            i_min = n - (N - M);

        if (i_max >= i_min)
        {
            double i_central = Math.round(n * M / N);
            if (i_central < i_min)
                i_central = i_min;
            if (i_central > i_max)
                i_central = i_max;

            double prob_central = hypergeometric_prob(N, M, n, i_central);
            result = prob_central;

            double prob = prob_central;
            for (double i = i_central - 1; i >= i_min; --i)
            {
                prob /= (n - i) / (i + 1) * (M - i) / (N - n - (M - i) + 1.0);
                result += prob;
            }

            prob = prob_central;
            for (double i = i_central + 1; i <= i_max; ++i)
            {
                prob *= (n - (i - 1)) / i * (M - (i - 1)) / (N - n - (M - (i - 1)) + 1.0);
                result += prob;
            }
        }
        return result;
    }

    private double M_to_Mplus1_factor(double N, double M, double n, double m) {
        double m_ratio = m / (M + 1);
        double tmp = m_ratio - (n - m) / (N - M);
        return 1.0 + tmp * (1.0 + m_ratio / (1.0 - m_ratio));
    }

    private double find_M_with_tail_prob(double N, double n, double m, double target_tail_prob, double M_initial,
                                         double M_guess, double M_bound) {
        double M_result = M_initial;

        double direction = (M_initial < M_bound) ? 1 : -1;

        double M_tolerance = 1;

        double cnt = 0;
        while (direction * (M_bound - M_result) > M_tolerance)
        {
            double prob = hypergeometric_prob(N, M_guess, n, m);

            double factor = M_to_Mplus1_factor(N, M_guess, n, m);

            double M_step;
            if (prob > target_tail_prob)
            {
                M_result = M_guess;
                if (Math.abs(factor - 1.0) < 0.00001)
                    M_step = Math.round((M_result + M_bound) / 2.0 - M_guess);
                else
                    M_step = Math.round(Math.log(target_tail_prob / prob) / Math.log(factor));
            }
            else
            {
                double cum_prob = hypergeometric_cumulative_prob(N, M_guess, n, m);
                double tail_prob = (direction > 0) ? cum_prob : 1.0 - cum_prob + prob;

                if (tail_prob > target_tail_prob)
                {
                    M_result = M_guess;
                    M_step = Math.round(Math.log(target_tail_prob / tail_prob) / Math.log(factor));
                }
                else
                {
                    M_bound = M_guess;
                    if (tail_prob < target_tail_prob / 1000.0)
                        M_step = Math.round((M_result + M_bound) / 2.0 - M_guess);
                    else
                        M_step = Math.round(Math.log(target_tail_prob / tail_prob) / Math.log(factor));
                }
            }

            double M_step_unadjusted = M_step;

            if (M_step == 0)
                M_step = direction;

            if (direction * (M_guess + M_step) >= direction * M_bound)
            {
                if (direction * (M_bound - M_guess) > 100)
                    M_step = Math.round(0.9 * M_bound + 0.1 * M_guess - M_guess);
                else
                    M_step = M_bound - direction - M_guess;
            }

            if (direction * (M_guess + M_step) <= direction * M_result)
            {
                if (direction * (M_guess - M_result) > 100)
                    M_step = Math.round(0.9 * M_result + 0.1 * M_guess - M_guess);
                else
                    M_step = M_result + direction - M_guess;
            }

            M_guess += M_step;

            ++cnt;
            if (cnt > 50)
                return 0.0; // undefined;
        }
        return M_result;
    }

    private double hypergeometric_approx_interval(double N, double n, double m, double confidence_level) {
        double num_sigma;
        double log_alpha = Math.log(1.0 - confidence_level);
        if (Math.abs(confidence_level - 0.90) < 0.000001)
            num_sigma = 1.64485;
        else if (Math.abs(confidence_level - 0.95) < 0.000001)
            num_sigma = 1.95996;
        else if (Math.abs(confidence_level - 0.98) < 0.000001)
            num_sigma = 2.32635;
        else if (Math.abs(confidence_level - 0.99) < 0.000001)
            num_sigma = 2.57583;
        else if (confidence_level <= 0.80)
            num_sigma = Math.exp(-0.6169356 - 0.537459 * log_alpha);
        else
            num_sigma = Math.exp(log_alpha * log_alpha / (log_alpha * log_alpha + 9.0) * (1.233943 - 0.0277894 * log_alpha));

        double p = m / n;
        double sigma = Math.sqrt((N - n) / (N - 1.0) * (p * (1.0 - p) + 1./n)  / n);

        return num_sigma * sigma;
    }


    /**
     *
     * @param N = Population Size
     * @param n = Sample Size
     * @param m = Number of Relevant Items in Sample
     * @param confidence_level
     * @return
     */

    private Map<String, Double> hypergeometric_confidence_interval(double N, double n, double m, double confidence_level)
    {
        confidence_level /= 100.0;

        double M_lower;
        double M_upper; //M_uppser;

        boolean flip = false;
        if (m > n - m)
        {
            flip = true;
            m = n - m;
        }


        double M_lowerbound = m - 1;
        double M_upperbound = N - (n - m) + 1;

        double M_mid = Math.round(m/n * N);

        if (M_mid <= M_lowerbound)
            M_mid = M_lowerbound + 1;
        if (M_mid >= M_upperbound)
            M_mid = M_upperbound - 1;

        M_lower = M_mid;
        M_upper = M_mid;

        double approx_interval = hypergeometric_approx_interval(N, n, m, confidence_level);

        double M_lower_guess = Math.round(M_mid - N * approx_interval);
        if (M_lower_guess <= M_lowerbound)
            M_lower_guess = M_lowerbound + 1;


        M_lower = find_M_with_tail_prob(N, n, m, (1.0 - confidence_level) / 2.0, M_lower, M_lower_guess, M_lowerbound);

        /*
        if (!typeof(M_lower) == 'number')
            return { };
         */

        double M_upper_guess = Math.round(M_lower + 2 * N * approx_interval);
        if (M_upper_guess >= M_upperbound)
            M_upper_guess = M_upperbound - 1;

        M_upper = find_M_with_tail_prob(N, n, m, (1.0 - confidence_level) / 2.0, M_upper, M_upper_guess, M_upperbound);

        if (flip)
        {
            double tmp = M_lower;
            M_lower = N - M_upper;
            M_upper = N - tmp;
        }

        //return { 'lower_bound': 100.0 * M_lower / N, 'upper_bound': 100.0 * M_upper / N };
        Map<String, Double> map = new HashMap<>();
        map.put("lower_bound", 100.0 * M_lower / N);
        map.put("upper_bound", 100.0 * M_upper / N);
        return map;
    }
}
