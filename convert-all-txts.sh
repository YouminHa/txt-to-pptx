PROGRAM_DIR=`dirname $0`

txt_files=( `ls *.txt` )
for f in ${txt_files[@]}
do
  out_file="${f%.*}.pptx"
  python ${PROGRAM_DIR}/txt-to-pptx.py -t ./template.pptx -i ${f} -o ${out_file}
done
