# CRISPR_Array_Generator
A python package for verifying and generating CRISPR Cas12 Arrays

CRISPR arrays that encode for multiple CRISPR gRNAs allow for multiplexed gene editing. Cas12 systems are best suited for this multiplexed targeting, since they possess the power to process gRNAs from an RNA transcript with no additional inputs. While this makes CRISPR Cas12 arrays a great tool for research labs to implement multiplexed gene targeting, designing these arrays can be tedious, as each gRNA within the array should be accompanied by separator, repeat, and annealing overhang sequences to optimize processing. With this, it is time-intensive and easy to make mistakes when designing these arrays by hand.

This tool can be used to check existing gRNAs or CRISPR arrays for common errors, and moreso, can be used to identify gRNAs from DNA sequences and auto-process gRNAs into ready-to-order array oligonucleotides. 

